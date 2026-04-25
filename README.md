# 招聘简历智能投递系统

基于 Claude AI 的简历解析、评分与管理系统。

---

## 项目结构

```
project/
├── app.py                  # 候选人投递页面（公开访问）
├── admin.py                # HR 管理后台（密码保护）
├── utils.py                # 核心工具库（解析/评分/导出）
├── candidates.json         # 候选人数据库（自动维护）
├── requirements.txt        # Python 依赖
├── .env.example            # 环境变量模板
├── .env                    # 你的真实配置（不要提交到 Git！）
├── .streamlit/
│   └── config.toml         # Streamlit 主题与上传限制配置
└── uploads/                # 简历/作品集存储目录（自动创建）
    └── 张三_8888/
        └── 20240601_103000/
            ├── resume.pdf
            └── portfolio.pdf
```

---

## 快速启动

### 第一步：安装依赖

```bash
pip install -r requirements.txt
```

### 第二步：配置环境变量

```bash
# 复制模板
cp .env.example .env

# 编辑 .env，填入你的 Anthropic API Key 和管理员密码
```

`.env` 文件内容示例：
```
ANTHROPIC_API_KEY=sk-ant-xxxxxxxx
ADMIN_PASSWORD=your_secure_password
```

### 第三步：启动两个页面

**投递页面**（供候选人使用，可公开）：
```bash
streamlit run app.py --server.port 8501
```

**HR 管理后台**（仅内部使用，密码保护）：
```bash
streamlit run admin.py --server.port 8502
```

> 两个命令在不同终端窗口分别运行。

---

## 使用说明

### 候选人端（app.py）
1. 填写基本信息（姓名、电话、邮箱等）
2. 上传简历（支持 TXT / DOCX / PDF）
3. 可选上传作品集
4. 点击提交 → AI 自动解析简历并评分
5. 页面显示 AI 识别结果供候选人核对

### HR 管理端（admin.py）
- **重复投递记录**：只展示多次投递的候选人，每次投递的时间和文件可单独下载
- **JD 综合评分**：7 维度 100 分评分汇总，可按岗位筛选，点开查看每维度理由
- **导出 Excel**：一键下载，包含两个 Sheet（候选人信息 + 重复投递记录），支持列筛选

---

## AI 功能说明

### 简历解析
调用 Claude API 从简历原文提取：编程技能、办公软件、AI工具熟练度、英语水平、实习经历、校园领导经历、个人特质等字段。

### 幻觉防控机制
- 提取的手机号和邮箱会和原文比对，格式不对或原文没有则自动清空并提示
- 技能列表只保留真实出现在原文的项，AI 凭空补充的会被过滤
- 评分由系统按维度累加，不用 AI 自报的 total（防止数字幻觉）
- 各维度分数被强制夹在 [0, 满分] 范围内

### 综合评分维度（满分 100 分）
| 维度 | 满分 |
|------|------|
| 个人性格与能力 | 30 |
| 编程技能 | 20 |
| 办公软件 | 10 |
| 英语能力 | 10 |
| AI 工具使用 | 10 |
| 相关实习经历 | 10 |
| 校园领导经历 | 10 |

---

## 注意事项

- `candidates.json` 和 `uploads/` 目录存储所有候选人数据，请定期备份
- `.env` 文件包含 API Key，**不要上传到任何公开代码仓库**
- 默认管理员密码在 `.env` 中设置，请使用强密码
- 上传文件大小限制默认 50MB，可在 `.streamlit/config.toml` 中调整
