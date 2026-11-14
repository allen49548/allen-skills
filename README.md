# 🧾 自动报销系统（reimbursement-skill）

> **一个仓库 = 一个技能。** 本仓库只收录**自动报销系统**这一个 Agent Skill：从飞书「物流部」混合群读取**付款凭证截图**，按司机与月份分类统计、幂等去重（绝不重复计算），写入司机报表「报销明细」sheet，每轮发两层通报。

兼容 WorkBuddy / Claude Code / 以及 40+ 支持 Agent Skills 开放标准的 Agent。`SKILL.md` 位于仓库根目录，`.claude-plugin/marketplace.json` 把它作为一个单技能插件分发。

> 仓库当前地址为 `allen49548/allen-skills`；若日后在 GitHub 上重命名为 `reimbursement-skill`，GitHub 会自动重定向旧地址，下面的命令无需改动。

---

## 🧾 它能做什么

- 拉取飞书群消息与图片资源，逐张视觉识别付款凭证截图（微信 / 支付宝账单、充电订单完成页等）
- 按**消息发送者**识别司机（司机弹性、只统计发了图片的司机、不纠结全员到齐）
- 按**司机 + 月份**定位报表，统一写入报表内唯一的「**报销明细**」sheet（只追加明细行，不改原有单元格）
- 四层去重（`message_id → 图片 SHA-256 → 感知哈希 → 订单号 / 业务键`），重复提交绝不重复计算
- 幂等台账（`state/ledger.sqlite`）+ 每次写报表前自动备份
- 两层通报：物流部群极简总览（金额 / 日期 / 安全驾驶调侃 / @报销人）/ 负责人私聊精简专业版
- 飞书确认任务（负责人可编辑，司机仅查看 / 评论），真实报销点完成后群内 @报销人回执

## 📁 目录结构

```
reimbursement-skill/
├── SKILL.md                 # 技能定义与铁律（必有，根目录）
├── config.example.json      # 脱敏配置模板（复制为 config.json 后填值）
├── requirements.txt         # 依赖（openpyxl / Pillow）+ lark-cli 前置说明
├── README.md
├── LICENSE                  # MIT
├── .claude-plugin/
│   └── marketplace.json     # 单技能插件清单：skills: ["./"]
├── .github/workflows/       # CI 校验（脚本编译 + frontmatter）
├── references/              # ocr-and-dedup / notification / report-layout
├── scripts/                 # 9 个核心脚本
│   ├── ledger.py            # 幂等台账、四层去重、四态、余额
│   ├── notify.py            # 两层通报发送（硬拦截司机私聊）
│   ├── driver_roster.py     # 弹性司机归属
│   ├── driver_copy.py       # 司机极简文案渲染
│   ├── owner_copy.py        # 负责人私聊精简版渲染
│   ├── group_task.py        # 报销确认任务 + 完成回执
│   ├── locate_report.py     # 报表定位
│   ├── reimb_detail.py      # 「报销明细」sheet 写入
│   └── monthly_summary.py   # 月度统计
├── state/
│   └── schema.sql           # 台账建表语句（首次运行前执行）
├── AGENTS.md                # 维护者 / 贡献规则
└── STATES.md                # 运行时状态说明
```

## 🔧 安装

**方式一：插件方式（推荐，自动跟随版本）**

```
/plugin marketplace add allen49548/allen-skills
/plugin install reimbursement-skill@reimbursement-skill
```

**方式二：手动拷贝**

把整个仓库目录拷进 Agent 读取 skills 的目录：

| 工具 | 所有项目可用 | 单项目可用 |
| --- | --- | --- |
| WorkBuddy | `~/.workbuddy/skills/<skill>/` | 项目内 `.workbuddy/skills/<skill>/` |
| Claude Code | `~/.claude/skills/<skill>/` | `.claude/skills/<skill>/` |
| Codex / Gemini CLI 等 | `~/.agents/skills/<skill>/` | `.agents/skills/<skill>/` |

## 🚀 使用前准备

1. **配置**：`cp config.example.json config.json`，把 `REPLACE_WITH_*` 占位符换成真实值（群 chat id、司机 open_id、`report_root` 本地绝对路径等）。`config.json` 已被 `.gitignore` 忽略，不会进库。
2. **建表**：首次运行前执行 `state/schema.sql` 初始化台账。
3. **装依赖**：`pip install -r requirements.txt`，并单独安装并登录 `lark-cli`（飞书命令行，非 pip）。
4. **跑一轮**：按 `SKILL.md` 流程调用脚本（`scripts/` 相对技能目录解析，脚本以自身位置定位 `config.json` 与 `state/`，因此目录整体放在哪都能跑）。

## 🔒 隐私与安全

`.gitignore` 排除所有真实隐私与产物，**绝不入库**：

- `config.json`（真实 chat id / open_id / 本地绝对路径）
- `lark-im-resources/`（飞书下载的真实付款凭证截图）
- `state/ledger.sqlite`（真实台账）、`state/backups/`（真实报表副本）
- `state/runs/`（运行日志含 chat_id / 任务 GUID / 金额）、`state/monthly_sent.json`

提交前务必 `git status` 复核，确保以上文件未被跟踪。

## 🛠 维护与贡献

规则见 [`AGENTS.md`](AGENTS.md)，运行时状态见 [`STATES.md`](STATES.md)。改动走 PR；CI 在 Ubuntu / macOS / Windows 校验脚本编译与 SKILL.md frontmatter。

## 📄 License

MIT License · 自由使用 / 修改 / 再分发
