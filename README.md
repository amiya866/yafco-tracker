# 财报跟踪平台 · 构建文档

全球锡/锌/铝/镍/铜/锂企业财报产量数据跟踪体系。本文档描述**从数据底稿目录到网站产出的完整构建链路**。

- 本地使用：双击 `index.html`（file:// 直开，无需服务器），密码 `yafco888`（防窥门槛，非加密）
- 外网版：https://wangziquan-del.github.io/yafco-tracker/ （GitHub Pages，仓库 public，同密码）
- 重建+上线一条龙：`bash deploy.sh`

## 〇、数据模型（品种 → 数据域 → 板块 → 公司）

```
财报数据库
├── 品种（一级）：锡 / 锌 / 铝 / 镍 / 铜 / 锂
│   ├── 1. 产量 ★核心 —— 资源端(矿/盐湖) / 冶炼端(锭/锂盐) / 品种中间层(镍=印尼中间品, 铝=三链)
│   ├── 2. 成本 —— 成本曲线(跨公司横比) / 公司成本序列(AISC/C1) / 财务质量备注
│   ├── 3. 资本开支与项目管线 —— 公司 capex / 在建扩产项目(Grasberg/AP60/多农/枧下窝)
│   ├── 4. 指引与预测 —— FY2026 指引 vs 年化 / 2027 展望(公司级+品种总量级)
│   ├── 5. 信息检索 —— 官方公告 / 行业媒体 / 市场传言(【传言】状态机)
│   ├── 6. 披露日历 —— 日期×公司×事件×状态
│   ├── 7. 更新日志 —— 数据血缘（何时改什么、来源）
│   └── 8. 口径元数据 —— 单位/折算/权益/财年映射
└── 跨品种（共享）：宏观资讯 / 供应强弱对比 / CRU 式区域总量参考
```

设计原则：**产量只是 8 个并列数据域之一**，新域（需求/库存/价格 TC 等待建）= 二级加分支 + 前端加区块，不动产量体系。当前缺口：capex 仅锌有、项目管线未独立成域、需求/库存/价格未建。

## 一、数据流总览（源头是爬取，Excel 是持久层）

```
【L0 源头 · 自动爬取为主】
  公司官网 IR / 交易所公告（ASX/SEC/港交所/巨潮）  ← 定时任务按 SOP 逐家核查
  阿拉丁 / 爱择（铝行业站，API+会员全文）
  微信公众号文章（用户丢链接，FetchURL 抓正文）
  抖音/小红书 MCP（供应端传言）
        │  核实、折算、标口径（禁止编造，未找到标「未找到」）
        ▼
【L1 底稿 · 持久层】
  各品种 Excel（产量/capex/更新日志）  ← 定时任务写入，写前备份
  data/news.json（信息速递条目）
  build_site.py 手工常量（日历/口径/综述观点/2027展望/成本曲线）
        │  python build_site.py
        ▼
【L2 构建产物】
  COMMODITIES 注册表 → extract_* → 统一 sections
    → 机制1 缺季拟合 → 机制2 指引进度 → 品种综述 → 2027展望 → 抽查核对 PASS/FAIL
  data/data.js + index.html 哈希文件名引用
        │
        ▼
【L3 呈现】
  本地 file:// ／ GitHub Pages（deploy.sh 推送）
```

要点：**Excel 不是数据源，是爬取结果的持久层**——数据的第一现场是官网公告和行业站，定时任务按各品种 SOP 自动核查写入；人工只处理口径判断和推算标注。


## 二、目录结构

### 平台目录 `D:\拷贝文件\E\永安\财报跟踪平台\`

```
build_site.py      # 数据管道（唯一构建脚本，含全部手工常量）
index.html         # 入口（密码门；资源引用为哈希文件名，构建时自动改写）
deploy.sh          # build + git commit + push（Pages 约 1 分钟生效）
assets/
  app.js           # 前端逻辑（纯 JS）；构建时生成 app.<hash>.js
  style.css        # 样式；构建时生成 style.<hash>.css
  echarts.min.js   # ECharts 本地化（CDN 仅兜底）
data/
  data.js          # 构建产物 window.SITE_DATA；构建时生成 data.<hash>.js
  news.json        # 信息速递源（手工+爬虫追加）
README.md          # 本文档
```

### 品种底稿目录（每个品种一份 Excel + 一份 SOP）

| 品种 | Excel 底稿 | SOP（更新规范/口径/披露节奏） | 板块结构 |
|---|---|---|---|
| 锡 | `永安\周报数据更新\进出口库存\海外主要公司产量.xlsx` | `永安\周报数据更新\AGENTS.md` | 矿山/精炼锡 |
| 锌 | `永安\锌\全球锌企季度产量梳理.xlsx` | `永安\锌\AGENTS.md` | 锌矿/锌锭冶炼 |
| 铝 | `永安\铝\全球铝企季度产量梳理.xlsx` | `永安\铝\AGENTS.md`（含阿拉丁/爱择抓取流程） | 铝土矿/氧化铝/电解铝 |
| 镍 | `永安\镍\全球镍企季度产量梳理.xlsx` | `永安\镍\AGENTS.md` | 印尼中间品/一级镍/镍铁NPI |
| 铜 | `永安\铜\全球铜企季度产量梳理.xlsx` | `永安\铜\AGENTS.md` | 铜矿(kt)/国内铜企(万吨) |
| 锂 | `永安\锂\全球锂企季度产量梳理.xlsx` | `永安\锂\AGENTS.md` | 锂资源(吨LCE)/锂盐冶炼(吨) |

锂的首次建库研究底稿在 `永安\锂\research\*.json`（逐公司来源+口径记录）。

## 三、构建管线 build_site.py

### 3.1 COMMODITIES 注册表

加品种 = 注册表加一条 + 写一个 `extract_*` 函数：

```python
"lithium": {"name": "锂", "excel": r"...xlsx", "extract": "extract_lithium",
            "calendar": LITHIUM_CALENDAR, "caliber_notes": LITHIUM_CALIBER_NOTES,
            "default_view": "quarter"}
```

抽取函数返回统一结构：`sections[]`（每板块 key/title/unit/quarters/years/companies/total）+ `costs`/`cost_curve`/`capex` + `changelog` + `overview`。前端按 sections 自动渲染，无需改前端。

### 3.2 数据机制

- **机制 1 缺季拟合**：年度有值、季度空缺时按季节分布/平均拟合，`est_q` 标记 → 前端斜体琥珀色显示
- **机制 2 指引进度**：`parse_guide_value` 解析 FY2026 指引文本（只读「；」前第一段，数值置前+单位），年化=Q1×4 或 H1×2，状态=超出/符合/不及
- **同比**：一律 Python 自算（Excel 公式不读缓存值，`data_only=False`）
- **总计行**：口径混杂直接求和，仅作量级参考（同比按两期均有值的公司集合）

### 3.3 手工维护常量（文件头区域）

| 常量 | 内容 | 更新时机 |
|---|---|---|
| `*_CALENDAR` | 各品种披露日历 | 披露节奏变化时 |
| `*_CALIBER_NOTES` | 口径说明（页底折叠区） | 口径变化时 |
| `REVIEW_COMMENTS` | 品种综述观点段（最新财报季总结+年度看法+α），**date 随内容改** | 每个财报节点 |
| `GLOBAL_SUPPLY_VIEW` | 跨品种供应端强弱对比（所有品种页共享） | 格局变化时 |
| `FY2027_OUTLOOK` | 2027 展望（按公司，标 官方/计划/平台推断；无依据的构建时自动按线性持平补填） | 每个财报节点 |
| 锂 cost_curve | 澳矿成本曲线（A$/dmt，FY26 指引） | 指引调整时 |

### 3.4 抽查核对

main() 末尾对六品种关键数值做 PASS/FAIL 核对（如 Alphamin 26Q2=5013、IGO 26Q2=48375），任一 FAIL 打印报警。**改 Excel 后必须跑 build 并确认全部 PASS。**

### 3.5 缓存破防

构建末尾把 `app.js`/`style.css`/`data.js` 复制为 `<名>.<内容哈希>.<后缀>` 并改写 index.html 引用（file:// 下 `?v=` 查询串不可靠，改文件名才彻底）。

## 四、前端结构（assets/app.js）

品种页区块顺序：**信息速递（最新 4 条，红黄绿 impact 标）→ 品种综述（自动统计+研究观点+跨品种对比）→ 产量板块（季度/年度切换）→ 公司卡片 → 2026 指引 vs 年化进度（含 2027 展望列+总量展望行）→ 成本与资本开支 → 披露日历/更新日志（默认折叠）→ 口径说明**。公司卡片/日历/日志均可点标题折叠。

工程细节：

- **ECharts 一律 SVG 渲染 + rAF 延迟初始化**（挂载前 init 会量成 0 宽图——2026-07-28 成本图不可见事故的根因，勿回退）
- 图表库加载失败时成本图降级为纯 HTML 条形图，产量图降级为提示+数据表
- 自诊断：JS 错误红色浮条 + 页脚「图表引擎」状态
- 密码门：SHA-256 校验，sessionStorage 记住

## 五、信息速递（data/news.json）

字段：`date / commodity(锡/锌/铝/镍/铜/锂/宏观) / category(供应/需求/政策/公司/宏观) / title / summary / source / url / impact(high红/mid黄/low绿)`。追加后重跑 build 即上线；单品种页限最新 4 条，总览页全量。

来源渠道：

- **公司官网/交易所**（财报节点定时任务顺手提炼）
- **阿拉丁 aladdiny.com**：铝供应端，免登录列表 API + 会员全文（流程见 `铝\AGENTS.md`，账号 yaqh）
- **爱择 azchina-cn.com**：铝独家/专题调研列表免登录（正文需会员，暂无账号）
- **抖音/小红书 MCP**：供应端传言——title 前缀【传言】、注明未证实、source 标「平台@作者」（小红书额度不稳定）
- **用户丢链接**（微信文章可直接 FetchURL）

## 六、定时任务体系

- **各品种财报节点**：锡 13 个、锌 8 个、铝/镍/铜/锂各 8 个年度循环 cron（多在 09:07），任务提示词指向对应品种 SOP
- **每周一 09:11 扫描**（id 51265eed）：印尼镍链动态、阿拉丁、爱择、抖音热点、突发公告 → news.json
- 财报节点任务同时负责：更新 Excel → 备份 → 更新日志 → 更新 REVIEW_COMMENTS/FY2027_OUTLOOK → build → 提炼新闻 →（外网）deploy.sh

## 七、部署

- 本地：双击 index.html
- 外网：`bash deploy.sh` = build + commit + push → GitHub Pages（仓库 wangziquan-del/yafco-tracker，public——免费版 Pages 不支持私有仓库；要私有化需 GitHub Pro 后改 visibility）

## 八、扩展新品种清单

1. 建品种目录 + Excel（列结构：公司|国家|项目/口径|23Q1…总计|同比|变化原因|FY2026指引|备注 + 更新日志 sheet）
2. 写品种 SOP（AGENTS.md）
3. build_site.py：COMMODITIES 加条目 + extract 函数 + 日历/口径常量 + 抽查核对
4. （可选）REVIEW_COMMENTS / FY2027_OUTLOOK 加品种条目
5. 跑 build 确认 PASS → deploy.sh 上线
6. 建年度循环 cron；根 AGENTS.md 登记品种段落
