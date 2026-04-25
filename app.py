"""
app.py — 合并版：候选人投递 + HR 管理后台
侧边栏底部有隐藏 HR 入口，输对密码后切换到管理后台。
"""

import os
from io import BytesIO
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
    load_candidates, save_candidates,
    save_to_excel, get_score_profile, SCORE_DIMENSIONS,
)

BASE            = Path(__file__).parent
CANDIDATES_PATH = BASE / "candidates.json"
EXPORT_PATH     = BASE / "candidates_export.xlsx"
ADMIN_PASSWORD  = os.environ.get("ADMIN_PASSWORD", "admin123")

st.set_page_config(page_title="云测Testin简历投递平台", layout="wide", page_icon="📋")

for key, default in {
    "view":               "candidate",
    "authenticated":      False,
    "pending_submission": None,
    "mismatch_warnings":  [],
}.items():
    if key not in st.session_state:
        st.session_state[key] = default


# ── 工具函数 ──────────────────────────────────────────────────────────────
def _do_submit(parsed_data: dict, resume_text: str):
    with st.spinner("正在提交，请稍候…"):
        score_result = jd_match_score(parsed_data, resume_text)
    parsed_data["jd_score"] = score_result
    candidates = load_candidates(CANDIDATES_PATH)
    candidates, is_update = update_candidate(parsed_data, candidates)
    save_candidates(candidates, CANDIDATES_PATH)
    if is_update:
        st.success("✅ 您已有投递记录，本次信息已成功更新！HR 将在审阅后与您联系。")
    else:
        st.success("✅ 简历提交成功！HR 将在审阅后与您联系。")


# ── 侧边栏 ────────────────────────────────────────────────────────────────
with st.sidebar:
    if st.session_state.view == "admin":
        st.markdown("### 管理员工具")
        if st.button("🚪 退出登录", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.view = "candidate"
            st.rerun()
        st.divider()
        st.markdown("### 📱 投递链接二维码")
        recruit_url = st.text_input("投递页面地址", value="https://testin-recruit.streamlit.app")
        if recruit_url.strip():
            try:
                import qrcode
                qr = qrcode.QRCode(box_size=6, border=2)
                qr.add_data(recruit_url.strip())
                qr.make(fit=True)
                img = qr.make_image(fill_color="#1F4E79", back_color="white")
                buf = BytesIO()
                img.save(buf, format="PNG")
                buf.seek(0)
                st.image(buf, caption="扫码直达投递页面", use_container_width=True)
                st.download_button("⬇️ 下载二维码", data=buf.getvalue(),
                                   file_name="recruit_qrcode.png", mime="image/png",
                                   use_container_width=True)
            except ImportError:
                st.warning("请先安装：pip install qrcode[pil]")
        st.divider()
        st.caption(f"登录时间：{datetime.now().strftime('%H:%M:%S')}")
    else:
        with st.expander("🔒 HR 入口", expanded=False):
            with st.form("login_form"):
                pwd = st.text_input("管理员密码", type="password",
                                    label_visibility="collapsed", placeholder="输入管理员密码")
                login_clicked = st.form_submit_button("登录", use_container_width=True)
            if login_clicked:
                if pwd == ADMIN_PASSWORD:
                    st.session_state.authenticated = True
                    st.session_state.view = "admin"
                    st.rerun()
                else:
                    st.error("密码错误")


# ════════════════════════════════════════════════════════════════════════════
# 候选人投递页面
# ════════════════════════════════════════════════════════════════════════════
def _show_candidate():
    st.title("📋 云测Testin简历投递平台")
    st.caption("请如实填写以下信息，系统将自动解析您的简历并保密存储。")

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
        strengths      = st.text_area("3~5 个优势/特点", height=100)
        resume_file    = st.file_uploader("上传简历 TXT / DOCX / PDF *", type=["txt","docx","pdf"])
        portfolio_file = st.file_uploader("上传个人作品集（可选）", type=["pdf","docx","zip"])
        submitted = st.form_submit_button("🚀 提交", use_container_width=True)

    if submitted:
        errors = []
        if not name.strip():          errors.append("姓名不能为空")
        if not validate_phone(phone): errors.append("手机号格式不正确（11位数字，以1开头）")
        if not validate_email(email): errors.append("邮箱格式不正确")
        if resume_file is None:       errors.append("请上传简历文件")
        if errors:
            for e in errors: st.error(e)
            st.stop()

        resume_text, extract_error = "", ""
        if resume_file:
            with st.spinner("正在处理简历…"):
                resume_text, extract_error = extract_resume_from_bytes(
                    resume_file.getvalue(), Path(resume_file.name).suffix)
        if extract_error:
            st.error(f"简历读取失败：{extract_error}")
            st.info("提示：扫描图片型简历无法提取文字，请上传可编辑的 DOCX 或 TXT 格式。")

        FILE_LIB   = BASE / "file_library"
        ts         = datetime.now().strftime("%Y%m%d_%H%M%S")
        submit_dir = FILE_LIB / phone.strip() / ts
        submit_dir.mkdir(parents=True, exist_ok=True)

        resume_path = portfolio_path = None
        if resume_file:
            resume_path = (submit_dir / f"resume{Path(resume_file.name).suffix}").resolve()
            resume_path.write_bytes(resume_file.getvalue())
        if portfolio_file:
            portfolio_path = (submit_dir / f"portfolio{Path(portfolio_file.name).suffix}").resolve()
            portfolio_path.write_bytes(portfolio_file.getvalue())

        parsed_data: dict = {}
        if resume_text:
            with st.spinner("正在提交，请稍候…"):
                parsed_data = parse_resume_text(resume_text)

        mismatches = []
        if resume_text:
            ai_phone = (parsed_data.get("phone") or "").strip()
            ai_email = (parsed_data.get("email") or "").strip()
            if ai_phone and ai_phone != phone.strip():
                mismatches.append(f"**电话**：您填写「{phone.strip()}」，简历中识别到「{ai_phone}」")
            if ai_email and ai_email != email.strip():
                mismatches.append(f"**邮箱**：您填写「{email.strip()}」，简历中识别到「{ai_email}」")

        parsed_data.update({
            "name": name.strip(), "gender": gender,
            "phone": phone.strip(), "email": email.strip(),
            "school": school.strip(), "major": major.strip(),
            "degree": degree, "target_position": target_position,
            "strengths": [s.strip() for s in strengths.splitlines() if s.strip()],
            "resume_file_path":    str(resume_path.resolve()) if resume_path else "",
            "portfolio_file_path": str(portfolio_path.resolve()) if portfolio_path else "",
            "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "resume_text":  resume_text,
        })

        if mismatches:
            st.session_state.pending_submission = parsed_data
            st.session_state.mismatch_warnings  = mismatches
        else:
            st.session_state.pending_submission = None
            st.session_state.mismatch_warnings  = []
            _do_submit(parsed_data, resume_text)

    if st.session_state.mismatch_warnings and st.session_state.pending_submission is not None:
        st.warning("⚠️ AI 识别到以下信息与您填写的不一致，请检查：")
        for msg in st.session_state.mismatch_warnings:
            st.markdown(f"- {msg}")
        st.caption("如确认填写无误（AI 识别偶尔有误差），可点击下方按钮继续提交。")
        if st.button("✅ 确定继续提交", use_container_width=True, type="primary"):
            pending = st.session_state.pending_submission
            st.session_state.pending_submission = None
            st.session_state.mismatch_warnings  = []
            _do_submit(pending, pending.get("resume_text", ""))

    st.divider()
    st.caption("🔒 您的简历和个人信息将被加密存储，仅授权 HR 人员可查看，不会对外公开。")


# ════════════════════════════════════════════════════════════════════════════
# HR 管理后台
# ════════════════════════════════════════════════════════════════════════════
def _show_admin():
    st.title("🛡️ 云测Testin HR管理后台")

    candidates = load_candidates(CANDIDATES_PATH)
    if not candidates:
        st.info("暂无候选人数据。")
        return

    st.caption(f"当前共有 **{len(candidates)}** 位候选人  ·  数据更新时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
    tab1, tab2, tab3 = st.tabs(["🔁 重复投递记录", "🎯 JD 综合评分", "📤 导出 Excel"])

    with tab1:
        st.subheader("重复投递记录")
        duplicates = [c for c in candidates if c.get("submission_history")]
        if not duplicates:
            st.success("✅ 目前没有重复投递记录。")
        else:
            for c in duplicates:
                cname, cphone = c.get("name","未知"), c.get("phone","")
                history = c.get("submission_history", [])
                with st.expander(f"👤 {cname}  |  📱 {cphone}  |  共投递 {len(history)+1} 次", expanded=True):
                    rows = [{"投递时间": c.get("last_updated","未知"), "版本":"最新",
                             "简历文件": c.get("resume_file_path",""), "作品集文件": c.get("portfolio_file_path","")}]
                    for h in reversed(history):
                        rows.append({"投递时间": h.get("timestamp","未知"), "版本":"历史",
                                     "简历文件": h.get("resume_file_path",""), "作品集文件": h.get("portfolio_file_path","")})
                    c1h,c2h,c3h,c4h = st.columns([2,1,3,3])
                    c1h.markdown("**投递时间**"); c2h.markdown("**版本**")
                    c3h.markdown("**简历文件**"); c4h.markdown("**作品集**")
                    st.divider()
                    for row in rows:
                        c1,c2,c3,c4 = st.columns([2,1,3,3])
                        c1.write(row["投递时间"]); c2.write(row["版本"])
                        rp = Path(row["简历文件"]) if row["简历文件"] else None
                        if rp and rp.exists():
                            with open(rp,"rb") as f:
                                c3.download_button(f"📄 {rp.name}", f, file_name=rp.name,
                                                   key=f"dl_r_{cname}_{row['投递时间']}")
                        else:
                            c3.write("—")
                        pp = Path(row["作品集文件"]) if row["作品集文件"] else None
                        if pp and pp.exists():
                            with open(pp,"rb") as f:
                                c4.download_button(f"📦 {pp.name}", f, file_name=pp.name,
                                                   key=f"dl_p_{cname}_{row['投递时间']}")
                        else:
                            c4.write("—")

    with tab2:
        st.subheader("JD 综合评分汇总")
        st.caption("评分由 AI 根据简历全文自动打出，满分 100 分。")
        scored = []
        for c in candidates:
            sd = c.get("jd_score", {})
            total = sd.get("total", None)
            scored.append({"_c": c, "姓名": c.get("name",""), "意向岗位": c.get("target_position",""),
                           "综合评分": total if total is not None else "未评分"})
        scored.sort(key=lambda x: x["综合评分"] if isinstance(x["综合评分"], int) else -1, reverse=True)
        positions = ["全部"] + sorted({c.get("target_position","") for c in candidates})
        sel_pos = st.selectbox("按意向岗位筛选", positions, key="score_filter")
        if sel_pos != "全部":
            scored = [s for s in scored if s["意向岗位"] == sel_pos]
        for idx, row in enumerate(scored):
            c  = row["_c"]
            sd = c.get("jd_score", {})
            total = row["综合评分"]
            label = f"**{idx+1}. {row['姓名']}**  |  {row['意向岗位']}  |  综合评分：{total}"
            if isinstance(total, int): label += " / 100"
            with st.expander(label):
                if sd.get("summary"): st.info(sd["summary"])
                profile = get_score_profile(sd.get("position", c.get("target_position","")))
                scores  = sd.get("scores", {})
                display_keys = list(scores.keys()) if scores else list(profile.keys())
                if not display_keys:
                    st.caption("暂无评分维度数据")
                else:
                    cols = st.columns(len(display_keys))
                    for i, dk in enumerate(display_keys):
                        dd   = scores.get(dk, {})
                        lbl  = dd.get("label") or profile.get(dk,{}).get("label", dk)
                        mmax = dd.get("max")   or profile.get(dk,{}).get("max", 10)
                        s    = dd.get("score", 0)
                        cols[i].metric(lbl, f"{s}/{mmax}",
                                       delta=f"{s-mmax//2:+d}" if isinstance(s,int) else None)
                        if dd.get("reason"): cols[i].caption(dd["reason"])
            st.divider()

    with tab3:
        st.subheader("导出候选人数据到 Excel")
        st.markdown("导出内容包含两个工作表：**候选人信息** 和 **重复投递记录**。")
        if st.button("📥 生成并下载 Excel", use_container_width=True):
            with st.spinner("正在生成 Excel 文件…"):
                save_to_excel(candidates, EXPORT_PATH)
            with open(EXPORT_PATH,"rb") as f:
                st.download_button("⬇️ 点击下载 candidates_export.xlsx", f,
                                   file_name="candidates_export.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
            st.success(f"Excel 已生成，共 {len(candidates)} 条记录。")


# ── 路由 ──────────────────────────────────────────────────────────────────
if st.session_state.view == "admin":
    _show_admin()
else:
    _show_candidate()
