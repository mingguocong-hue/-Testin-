"""
app.py — 用户简历投递页面
候选人只看到提交成功/失败，AI识别结果和评分均不展示给候选人。
"""

import json
from pathlib import Path
from datetime import datetime

import streamlit as st

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from utils import (
    validate_phone, validate_email,
    extract_resume_from_bytes,
    parse_resume_text,
    update_candidate, jd_match_score,
)

BASE = Path(__file__).parent

st.set_page_config(page_title="云测Testin简历投递平台", layout="wide", page_icon="📋")
st.title("📋 云测Testin简历投递平台")
st.caption("请如实填写以下信息，系统将自动解析您的简历并保密存储。")

# ────────────────────────────────────────────────────────────────────────────
# 投递表单
# ────────────────────────────────────────────────────────────────────────────
# ── Session State 初始化（信息不一致二次确认）────────────────────────────
if "pending_submission" not in st.session_state:
    st.session_state.pending_submission = None
if "mismatch_warnings" not in st.session_state:
    st.session_state.mismatch_warnings = []

with st.form("resume_form"):
    col1, col2 = st.columns(2)
    with col1:
        name   = st.text_input("姓名 *")
        phone  = st.text_input("电话 *")
        school = st.text_input("毕业院校 *")
        degree = st.selectbox("学历", ["本科在读", "硕士在读", "博士在读", "其他"])
    with col2:
        gender          = st.selectbox("性别", ["女", "男"])
        email           = st.text_input("邮箱 *")
        major           = st.text_input("专业")
        target_position = st.selectbox(
            "意向岗位", ["管培生", "运营岗位", "AI测试员", "后台技术管理员", "其他"]
        )

    strengths      = st.text_area("3~5 个优势/特点（每行一个）", height=100)
    resume_file    = st.file_uploader("上传简历 TXT / DOCX / PDF *", type=["txt", "docx", "pdf"])
    portfolio_file = st.file_uploader("上传个人作品集（可选）", type=["pdf", "docx", "zip"])

    submitted = st.form_submit_button("🚀 提交", use_container_width=True)

# ────────────────────────────────────────────────────────────────────────────
# 提交处理
# ────────────────────────────────────────────────────────────────────────────
if submitted:
    # ── 前端校验 ──────────────────────────────────────────────────────────
    errors = []
    if not name.strip():
        errors.append("姓名不能为空")
    if not validate_phone(phone):
        errors.append("手机号格式不正确（11位数字，以1开头）")
    if not validate_email(email):
        errors.append("邮箱格式不正确")
    if resume_file is None:
        errors.append("请上传简历文件")

    if errors:
        for e in errors:
            st.error(e)
        st.stop()

    # ── 从上传字节流提取简历文本 ────────────────────────────────────────────
    resume_text   = ""
    extract_error = ""
    if resume_file is not None:
        with st.spinner("正在处理简历…"):
            file_bytes  = resume_file.getvalue()
            file_suffix = Path(resume_file.name).suffix
            resume_text, extract_error = extract_resume_from_bytes(file_bytes, file_suffix)

    # 提取失败给候选人看错误原因（这是操作指引，不是评分信息）
    if extract_error:
        st.error(f"简历读取失败：{extract_error}")
        st.info("提示：扫描图片型简历无法提取文字，请上传可编辑的 DOCX 或 TXT 格式。")

    # ── 文件存储：file_library/<手机号>/<时间戳>/ ─────────────────────────
    FILE_LIB   = BASE / "file_library"
    ts         = datetime.now().strftime("%Y%m%d_%H%M%S")
    submit_dir = FILE_LIB / phone.strip() / ts
    submit_dir.mkdir(parents=True, exist_ok=True)

    resume_path    = None
    portfolio_path = None

    if resume_file is not None:
        resume_path = (submit_dir / f"resume{Path(resume_file.name).suffix}").resolve()
        resume_path.write_bytes(resume_file.getvalue())

    if portfolio_file is not None:
        portfolio_path = (submit_dir / f"portfolio{Path(portfolio_file.name).suffix}").resolve()
        portfolio_path.write_bytes(portfolio_file.getvalue())

    # ── AI 简历解析（后台静默处理，候选人不可见）──────────────────────────
    parsed_data: dict = {}
    if resume_text:
        with st.spinner("正在提交，请稍候…"):
            parsed_data = parse_resume_text(resume_text)

    # ── 信息一致性检查：比对 AI 识别 vs 表单填写 ──────────────────────────
    mismatches = []
    if resume_text:  # 只有成功提取简历文本时才做比对
        ai_name  = (parsed_data.get("name")  or "").strip()
        ai_phone = (parsed_data.get("phone") or "").strip()
        ai_email = (parsed_data.get("email") or "").strip()
        if ai_name  and ai_name  != name.strip():
            mismatches.append(f"**姓名**：您填写的是「{name.strip()}」，简历中识别到的是「{ai_name}」")
        if ai_phone and ai_phone != phone.strip():
            mismatches.append(f"**电话**：您填写的是「{phone.strip()}」，简历中识别到的是「{ai_phone}」")
        if ai_email and ai_email != email.strip():
            mismatches.append(f"**邮箱**：您填写的是「{email.strip()}」，简历中识别到的是「{ai_email}」")

    # ── 合并数据：表单字段优先，覆盖 AI 提取结果 ──────────────────────────
    parsed_data.update({
        "name":                name.strip(),
        "gender":              gender,
        "phone":               phone.strip(),
        "email":               email.strip(),
        "school":              school.strip(),
        "major":               major.strip(),
        "degree":              degree,
        "target_position":     target_position,
        "strengths":           [s.strip() for s in strengths.splitlines() if s.strip()],
        "resume_file_path":    str(resume_path.resolve()) if resume_path else "",
        "portfolio_file_path": str(portfolio_path.resolve()) if portfolio_path else "",
        "last_updated":        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "resume_text":         resume_text,
    })

    if mismatches:
        # 有不一致：暂存数据，等候选人二次确认
        st.session_state.pending_submission = parsed_data
        st.session_state.mismatch_warnings  = mismatches
    else:
        # 无不一致：直接完成提交
        st.session_state.pending_submission = None
        st.session_state.mismatch_warnings  = []
        _do_submit(parsed_data, resume_text, BASE)

# ────────────────────────────────────────────────────────────────────────────
# 完成提交的公共函数
# ────────────────────────────────────────────────────────────────────────────
def _do_submit(parsed_data: dict, resume_text: str, base: Path):
    """AI评分 → 写入 candidates.json → 展示结果。"""
    with st.spinner("正在提交，请稍候…"):
        score_result = jd_match_score(parsed_data, resume_text)
    parsed_data["jd_score"] = score_result

    candidates_path = base / "candidates.json"
    try:
        candidates = json.loads(candidates_path.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        candidates = []

    candidates, is_update = update_candidate(parsed_data, candidates)
    candidates_path.write_text(
        json.dumps(candidates, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    if is_update:
        st.success("✅ 您已有投递记录，本次信息已成功更新！HR 将在审阅后与您联系。")
    else:
        st.success("✅ 简历提交成功！HR 将在审阅后与您联系。")

# ────────────────────────────────────────────────────────────────────────────
# 信息不一致二次确认区域
# ────────────────────────────────────────────────────────────────────────────
if st.session_state.mismatch_warnings and st.session_state.pending_submission is not None:
    st.warning("⚠️ AI 识别到以下信息与您填写的不一致，请检查是否填写正确：")
    for msg in st.session_state.mismatch_warnings:
        st.markdown(f"- {msg}")
    st.caption("如果您确认填写信息无误（AI 识别偶尔会有误差），可点击下方按钮继续提交。")
    if st.button("✅ 确定继续提交", use_container_width=True, type="primary"):
        pending = st.session_state.pending_submission
        resume_text_pending = pending.get("resume_text", "")
        st.session_state.pending_submission = None
        st.session_state.mismatch_warnings  = []
        _do_submit(pending, resume_text_pending, BASE)

# ────────────────────────────────────────────────────────────────────────────
# 页脚
# ────────────────────────────────────────────────────────────────────────────
st.divider()
st.caption(
    "🔒 您的简历和个人信息将被加密存储，仅授权 HR 人员可查看，不会对外公开。"
)
