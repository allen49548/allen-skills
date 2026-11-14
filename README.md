# 🧾 自动报销系统（Auto Reimbursement）

> 从飞书「物流部」混合群读取**付款凭证截图**，按司机与月份分类统计报销，幂等去重（绝不重复计算），累加写入司机报表，每轮发两层通报。

一个 WorkBuddy / Claude Code / 以及 40+ 支持 Agent Skills 标准的 Agent 可直接加载的结构化技能。

---

## 📋 它能做什么

- 拉取飞书群消息与图片资源，逐张视觉识别付款凭证截图（微信 / 支付宝账单、充电订单完成页等）
- 按**消息发送者**识别司机（司机弹性、只统计发了图片的司机、不纠结全员到齐）
- 按**司机 + 月份**定位报表并做加法累加（zip 层精确写单元格，不动公式 / 格式）
- 四层去重（`message_id → 图片 SHA-256 → 感知哈希 → 订单号 / 业务键`），重复提交绝不重复计算
- 幂等台账（`state/ledger.sqlite`）+ 每次写报表前自动备份
- 两层通报：物流部群一份极简总览（金额 / 日期 / 安全驾驶调侃）/ 负责人私聊专业完整版
- 飞书确认任务（负责人可编辑，司机仅查看 / 评论）

## 📦 目录结构

```
自动报销系统/
├── SKILL.md                 # 技能定义与铁律（必有）
├── config.example.json      # 脱敏配置模板（复制为 config.json 后填值）
├── requirements.txt         # 依赖（openpyxl / Pillow）+ lark-cli 前置说明
├── references/              # 规则文档
│   ├── ocr-and-dedup.md
│   ├── notification.md
│   └── report-layout.md
├── scripts/                 # 7 个核心脚本
│   ├── ledger.py            # 幂等台账、四层去重、四态、余额
│   ├── notify.py            # 两层通报发送（硬拦截司机私聊）
│   ├── driver_roster.py     # 弹性司机归属
│   ├── driver_copy.py       # 司机极简文案渲染
│   ├── group_task.py        # 报销确认任务
│   ├── locate_report.py     # 报表定位
│   └── monthly_summary.py   # 月度统计
├── state/
│   └── schema.sql           # 台账建表语句（首次运行前执行）
├── AGENTS.md                # 维护者 / 贡献规则
├── STATES.md                # 运行时状态说明
├── LICENSE                  # MIT
├── .claude-plugin/          # 插件分发清单
└── .github/workflows/       # CI 校验
```

> 仓库遵循 Agent Skills 开放标准。**单个技能**时 `SKILL.md` 放在仓库根（WorkBuddy 等也正是这样加载）；要扩成多技能合集时，把本技能移入子目录（如 `reimbursement/`）并在 `.claude-plugin/marketplace.json` 的 `skills` 里登记即可。

## 🔧 安装

把整个仓库（含 `references/`、`scripts/`、`state/`）放进 Agent 读取的 skills 目录：

| 工具 | 所有项目可用 | 单项目可用 |
| --- | --- | --- |
| WorkBuddy | `~/.workbuddy/skills/<skill>/` | 项目内 `.workbuddy/skills/<skill>/` |
| Claude Code | `~/.claude/skills/<skill>/` | `.claude/skills/<skill>/` |
| Codex / Gemini CLI 等 | `~/.agents/skills/<skill>/` | `.agents/skills/<skill>/` |

或直接用仓库地址安装（把 `<repo>` 换成你的 GitHub 仓库）：

```
帮我安装这个 skill：https://github.com/<you>/<repo>/tree/main
```

## 🚀 使用前准备

1. **配置**：`cp config.example.json config.json`，把 `REPLACE_WITH_*` 占位符换成真实值（群 chat id、司机 open_id、`report_root` 本地绝对路径等）。`config.json` 已被 `.gitignore` 忽略，不会进库。
2. **建表**：首次运行前执行 `state/schema.sql` 初始化台账。
3. **装依赖**：`pip install -r requirements.txt`，并单独安装并登录 `lark-cli`（飞书命令行，非 pip）。
4. **跑一轮**：在 WorkBuddy 里 `@skill:"自动报销系统" 运行`，或按 SKILL.md 流程手动调用脚本。

## 🔒 隐私与安全

`.gitignore` 已排除所有真实隐私与产物，**绝不入库**：

- `config.json`（真实 chat id / open_id / 本地绝对路径）
- `lark-im-resources/`（飞书下载的真实付款凭证截图）
- `state/ledger.sqlite`（真实台账）、`state/backups/`（真实报表副本）
- `state/runs/`（运行日志含 chat_id / 任务 GUID / 金额）、`state/monthly_sent.json`

提交前务必 `git status` 复核，确保以上文件未被跟踪。

## 🛠 维护与贡献

规则见 [AGENTS.md](AGENTS.md)。运行时状态说明见 [STATES.md](STATES.md)。改动走 PR；CI 在 Ubuntu / macOS / Windows 校验脚本编译与 SKILL.md frontmatter。

## 📄 License

MIT License · 自由使用 / 修改 / 再分发
