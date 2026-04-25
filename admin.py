"""
admin.py — 管理员后台
功能：密码验证 → 重复投递记录查看 → JD综合评分汇总 → 一键导出Excel
"""

import json
from pathlib import Path
from datetime import datetime

import streamlit as st

from utils import save_to_excel, SCORE_DIMENSIONS, get_score_profile

BASE = Path(__file__).parent
CANDIDATES_PATH = BASE / "candidates.json"
EXPORT_PATH     = BASE / "candidates_export.xlsx"

# ────────────────────────────────────────────────────────────────────────────
# 页面配置
# ────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="云测Testin HR管理后台", layout="wide", page_icon="🛡️")
st.title("🛡️ 云测Testin HR管理后台")

# ────────────────────────────────────────────────────────────────────────────
# 密码验证（Session State 持久化）
# ────────────────────────────────────────────────────────────────────────────
ADMIN_PASSWORD = "admin123"   # 生产环境请改为环境变量

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    with st.form("login_form"):
        pwd = st.text_input("请输入管理员密码", type="password")
        login_clicked = st.form_submit_button("登录", use_container_width=True)
    if login_clicked:
        if pwd == ADMIN_PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("密码错误，拒绝访问")
    st.stop()

# ────────────────────────────────────────────────────────────────────────────
# 读取候选人数据
# ────────────────────────────────────────────────────────────────────────────
candidates = []
if CANDIDATES_PATH.exists():
    try:
        candidates = json.loads(CANDIDATES_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        st.error("candidates.json 格式错误，请检查文件")
        st.stop()

if not candidates:
    st.info("暂无候选人数据。")
    st.stop()

st.caption(f"当前共有 **{len(candidates)}** 位候选人  ·  数据更新时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}")

# ────────────────────────────────────────────────────────────────────────────
# Tab 布局
# ────────────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["🔁 重复投递记录", "🎯 JD 综合评分", "📤 导出 Excel"])

# ════════════════════════════════════════════════════════════════════════════
# Tab 1：重复投递记录
# 只展示有 submission_history（即曾经多次投递）的候选人
# ════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("重复投递记录")
    st.caption("以下候选人曾多次投递，每行对应一次投递，可点击链接查看对应文件。")

    duplicates = [c for c in candidates if c.get("submission_history")]

    if not duplicates:
        st.success("✅ 目前没有重复投递记录。")
    else:
        for c in duplicates:
            name  = c.get("name", "未知")
            phone = c.get("phone", "")
            history = c.get("submission_history", [])

            with st.expander(
                f"👤 {name}  |  📱 {phone}  |  共投递 {len(history) + 1} 次",
                expanded=True
            ):
                # 构建每次投递的行数据（最新在最上面）
                rows = []

                # 最新记录排第一
                rows.append({
                    "投递时间":   c.get("last_updated", "未知"),
                    "版本":       "最新",
                    "简历文件":   c.get("resume_file_path", ""),
                    "作品集文件": c.get("portfolio_file_path", ""),
                })

                # 历史记录（晚 → 早）
                for h in reversed(history):
                    rows.append({
                        "投递时间":   h.get("timestamp", "未知"),
                        "版本":       "历史",
                        "简历文件":   h.get("resume_file_path", ""),
                        "作品集文件": h.get("portfolio_file_path", ""),
                    })

                # 展示为表格，并对每行提供文件下载按钮
                col_time, col_ver, col_resume, col_port = st.columns([2, 1, 3, 3])
                col_time.markdown("**投递时间**")
                col_ver.markdown("**版本**")
                col_resume.markdown("**简历文件**")
                col_port.markdown("**作品集**")
                st.divider()

                for row in rows:
                    c1, c2, c3, c4 = st.columns([2, 1, 3, 3])
                    c1.write(row["投递时间"])
                    c2.write(row["版本"])

                    # 简历下载按钮
                    rp = Path(row["简历文件"]) if row["简历文件"] else None
                    if rp and rp.exists():
                        with open(rp, "rb") as f:
                            c3.download_button(
                                label=f"📄 {rp.name}",
                                data=f,
                                file_name=rp.name,
                                key=f"dl_resume_{name}_{row['投递时间']}",
                            )
                    else:
                        c3.write("—")

                    # 作品集下载按钮
                    pp = Path(row["作品集文件"]) if row["作品集文件"] else None
                    if pp and pp.exists():
                        with open(pp, "rb") as f:
                            c4.download_button(
                                label=f"📦 {pp.name}",
                                data=f,
                                file_name=pp.name,
                                key=f"dl_port_{name}_{row['投递时间']}",
                            )
                    else:
                        c4.write("—")

# ════════════════════════════════════════════════════════════════════════════
# Tab 2：JD 综合评分汇总
# ════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("JD 综合评分汇总")
    st.caption(
        "评分由 AI 根据简历全文自动打出，共 7 个维度满分 100 分。"
        "点击候选人姓名可查看各维度明细。"
    )

    # 汇总表
    scored = []
    for c in candidates:
        score_data = c.get("jd_score", {})
        total = score_data.get("total", None)
        scored.append({
            "_candidate": c,
            "姓名":       c.get("name", ""),
            "意向岗位":   c.get("target_position", ""),
            "综合评分":   total if total is not None else "未评分",
            "评分摘要":   score_data.get("summary", ""),
        })

    # 按分数排序（未评分排末尾）
    scored.sort(key=lambda x: x["综合评分"] if isinstance(x["综合评分"], int) else -1, reverse=True)

    # 筛选器
    positions = ["全部"] + sorted({c.get("target_position", "") for c in candidates})
    sel_pos = st.selectbox("按意向岗位筛选", positions, key="score_filter")
    if sel_pos != "全部":
        scored = [s for s in scored if s["意向岗位"] == sel_pos]

    # 展示每人评分行
    for idx, row in enumerate(scored):
        c = row["_candidate"]
        score_data = c.get("jd_score", {})
        total      = row["综合评分"]
        label_str  = f"**{idx+1}. {row['姓名']}**  |  {row['意向岗位']}  |  综合评分：{total}"
        if isinstance(total, int):
            label_str += " / 100"

        with st.expander(label_str):
            # 总评
            if score_data.get("summary"):
                st.info(score_data["summary"])

            # 动态获取该候选人岗位对应的维度配置
            position   = score_data.get("position", c.get("target_position", ""))
            profile    = get_score_profile(position)
            scores     = score_data.get("scores", {})

            # 优先用 jd_score.scores 中实际存在的维度，兜底用 profile
            display_keys = list(scores.keys()) if scores else list(profile.keys())
            if not display_keys:
                st.caption("暂无评分维度数据")
            else:
                cols = st.columns(len(display_keys))
                for i, dim_key in enumerate(display_keys):
                    dim_data = scores.get(dim_key, {})
                    # label/max 优先从 scores 里取（AI 评分时已写入），兜底从 profile 取
                    label  = dim_data.get("label") or profile.get(dim_key, {}).get("label", dim_key)
                    mmax   = dim_data.get("max")   or profile.get(dim_key, {}).get("max", 10)
                    s      = dim_data.get("score", 0)
                    reason = dim_data.get("reason", "")
                    cols[i].metric(
                        label=label,
                        value=f"{s}/{mmax}",
                        delta=f"{s - mmax//2:+d}" if isinstance(s, int) else None,
                    )
                    if reason:
                        cols[i].caption(reason)

        st.divider()

# ════════════════════════════════════════════════════════════════════════════
# Tab 3：导出 Excel
# ════════════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("导出候选人数据到 Excel")
    st.markdown(
        """
        导出内容包含两个工作表：
        - **候选人信息**：全量候选人，支持按岗位 / 学历 / 性别等筛选，简历/作品集列为可点击超链接
        - **重复投递记录**：有多次投递记录的候选人，展示每次投递时间与文件链接
        """
    )

    if st.button("📥 生成并下载 Excel", use_container_width=True):
        with st.spinner("正在生成 Excel 文件…"):
            save_to_excel(candidates, EXPORT_PATH)

        with open(EXPORT_PATH, "rb") as f:
            st.download_button(
                label="⬇️ 点击下载 candidates_export.xlsx",
                data=f,
                file_name="candidates_export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        st.success(f"Excel 已生成，共 {len(candidates)} 条记录。")

# ────────────────────────────────────────────────────────────────────────────
# 侧边栏：登出
# ────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 管理员工具")
    if st.button("🚪 退出登录"):
        st.session_state.authenticated = False
        st.rerun()
    st.divider()
    st.caption(f"登录时间：{datetime.now().strftime('%H:%M:%S')}")
