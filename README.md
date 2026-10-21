# 每日打卡记录系统

> Daily Check-in System — 把每一天的自律，变成一条看得见的收益曲线。
> A self-hosted habit tracker that turns daily check-ins into a benefit ledger, built on Flask + Excel.

单人使用的习惯打卡与个人账本合一的 Web 小系统：多级分类管理打卡项，勾选/多次/数值三种记录方式，所有行为按“收益”记账，自动生成每日收益链、自然月统计、周报月报与收益曲线。数据就是一个 Excel 文件——**透明、可带走、不锁死**。

![Python](https://img.shields.io/badge/Python-3.10%2B-blue) ![Flask](https://img.shields.io/badge/Flask-3.x-black) ![openpyxl](https://img.shields.io/badge/openpyxl-3.1%2B-green) ![License](https://img.shields.io/badge/License-MIT-yellow)

---

## 为什么做这个

市面上的打卡 App 面临同一个取舍：好用的不自由（数据在别人服务器），自由的不好用（自己要搭）。这个项目的立场是——**数据必须在自己手里，双击就能打开核对**，所以选了 Excel 做存储；同时把“打卡”升级成“记账”：每完成一个习惯获得收益，犯一个坏习惯被扣款，收益链一天天累积，让抽象的“坚持”变成具体的数字。

## 功能特性

**记录**
- 无限级分类树：养生 / 手臂 / 体态 / 核心 / 专注 / 腿部……随意嵌套
- 四种项目类型：分类、勾选（做了/没做）、多次项（如八步金刚功每日 4 次，达标才结算）、数值项（如专注 2.5 小时，支持负数）
- 过去日期补录：漏打哪天补哪天，当日收益链自动重算

**账目**
- 收益链：当日新增 = Σ勾选收益 + Σ多次达标收益 + Σ数值；累计 = 上日累计 + 当日新增，初始值为常量
- 每日收益表 + 收益曲线（手绘 SVG）
- 自然月统计（按已过天数算完成率，不与 31 天硬比）、月份切换、环比
- 周报 / 月报 / 月末余额预测（按当前日均外推）
- 热力图、周/月完成率趋势、项目贡献图

**数据安全**
- 每次保存前自动备份（每日一份、保留 30 份）、临时文件原子替换、Excel 占用检测
- 软删除：归档项目历史与收益分毫不动；独创「归档记录」冷表，归档即搬家，主表永远小而快
- 一键导出备份 zip（当前数据 + 全部历史备份 + 恢复说明）
- 数据体检：发现手工改表造成的“孤儿列”会主动提醒

**体验**
- 原生 HTML/JS/CSS 前端，零 CDN、零构建、断网可用
- 历史明细折叠 / 时间范围筛选 / 分批加载，偏好本地记忆
- 归档箱支持重命名回收名称、勾选项与多次项可互转且历史数据自动迁移

## 快速开始

环境要求：Python 3.10+（开发使用 3.13）

```bash
git clone https://github.com/<your-name>/Daily.git
cd Daily
pip install -r requirements.txt
python app.py
```

启动后访问 <http://127.0.0.1:5000>。首次启动会自动创建 `data/打卡记录.xlsx`，无需任何初始化。

Windows 用户也可以直接双击 `start.bat`（自带防重复启动、3 秒后自动打开浏览器）。

## 数据文件说明

全部数据就是 `data/` 下的一个 Excel 文件，三个 sheet：

| Sheet | 结构 | 说明 |
|---|---|---|
| `活动结构` | ID / 名称 / 类型 / 父级ID / 单位 / 收益 / 目标次数 / 状态 | 邻接表存树，支持无限层级；`状态` = active / archived |
| `打卡记录` | 行 = 日期（YYYY-MM-DD），列 = 叶子项 | 勾选存 1/0；多次项存当日次数（0~目标）；数值项存具体值 |
| `每日收益` | 日期 / 当日新增 / 累计收益 | 由打卡记录全量派生，可随时重算 |

`归档记录`（冷表）在首次归档时自动创建，结构为 项目 / 日期 / 值，只存有记录的日期。

**核心原则：单一账本**。收益链永远从打卡记录派生，不单独维护——任何补录、改价、归档、删除之后，全局重算一次，账自动平。

## 项目结构

```
Daily/
├── app.py               # 全部后端：数据层 + 17 个 API + 同步引擎
├── templates/
│   └── index.html       # 全部前端：单页 + 原生 JS（图表为手绘 SVG）
├── data/                # 运行时生成：打卡记录.xlsx + backup/
├── test_sync.py         # 单元测试（同步引擎 / 收益链 / 统计 / 迁移）
├── smoke_test.py        # 接口联调（先启动服务再跑）
├── start.bat            # Windows 一键启动
└── requirements.txt
```

## API 一览

| 方法 | 路径 | 用途 |
|---|---|---|
| GET | `/api/state` | 今日状态（树 / 进度 / 收益 / 连续打卡 / 体检提示） |
| GET | `/api/history` | 历史：明细 / 月度统计 / 周报月报 / 热力图 / 趋势 / 贡献 |
| POST | `/api/toggle` | 勾选 / 取消勾选 |
| POST | `/api/value` | 数值项录入 |
| POST | `/api/count` | 多次项 +1/-1 |
| POST | `/api/day` | 历史某日整体补录 |
| POST | `/api/item/add` · `/update` · `/rename` | 新增 / 修改（支持类型互转）/ 重命名 |
| POST | `/api/item/remove` · `/restore` · `/purge` | 归档 / 恢复 / 永久删除 |
| GET/POST | `/api/settings` | 收益目标设置 |
| POST | `/api/backup` | 立即备份 |
| GET | `/api/backups` | 备份列表 |
| GET | `/api/export/zip` | 下载备份包 |

## 运行测试

```bash
python test_sync.py    # 单元测试，约 20 组
python app.py          # 另开一个终端先启动服务
python smoke_test.py   # 接口联调，覆盖全部 API 与边界
```

## 设计决策记录

1. **Excel 而不是 SQLite**：账本的透明性优先——双击可开、可用计算器复核、导出即备份。代价是每次读写全量解析文件，靠“归档冷热分离”把主表控制在几百行级；若未来出现复杂多维分析需求，数据可整体导入 SQLite，随时可迁。
2. **单文件后端**：个人系统没有多人协作与部署复杂度，一个 `app.py` 读到底、改到底，比拆分层更省心。
3. **软删除而非物理删除**：删除一个项目不等于抹掉它的历史——归档保留证据与收益，永久删除才重算账本。

> 📸 **截图占位**：把运行截图放在 `docs/` 下，然后取消下面两行注释
> `![今日打卡](docs/today.png)`
> `![历史统计](docs/history.png)`

## License

[MIT](LICENSE) © 2026
