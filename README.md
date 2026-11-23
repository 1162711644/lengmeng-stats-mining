# 冷檬数据统计需求挖掘 · Skill

> 给一张干净的表，先帮你挖出"这表能算什么"，再写一句能直接复制的统计指令，最后用**公式**把结果落到新增 sheet 里——原表不动，数据改了自动重算。

一个面向豆包 / WorkBuddy / Coze 等 Agent 运行时的中文 Skill，自包含、无外部依赖（仅需 Python + openpyxl）。

## 为什么做这个

拿到一张整理干净的业务表，最常见的卡点不是"不会算"，而是"不知道该算什么"。
这个 Skill 把过程拆成两段：

1. **挖掘**：读表 → 用大白话列出值得统计的业务数据 → 一句话引导你选 → 写出可整句复制的【做统计】指令。
2. **落地**：你把这句指令发回来，它在原表副本里新增 sheet，用 `SUMIFS / COUNTIFS / AVERAGEIFS / RANK / INDEX-MATCH` 算出结果。

只做统计计算（求和、平均、计数、分组、排名、占比、比值、对比），不做出图、不做分析结论——那是下游技能的事。

## 特性

- **单表 / 多表自动分流**：1 张表走单表版提示词；2 个文件或多 sheet 走多表版，自动帮你找关联键。
- **只引导不拍板**：只问"你最想看哪一类"，给方向选项，不替你决定业务重点。
- **指令即用**：输出的是一整句话，复制粘贴就能执行，不是填空模板；同时给"够用版"和"多看一点版"。
- **公式驱动**：结果 sheet 里每个指标都是公式，不是算好的死数字；原表增删数据后自动重算。
- **原表零改动**：只写副本，原 sheet 名称、顺序、数据都不变。
- **带口径备注**：每张结果 sheet 底部写明数据源、表头行、数据区、分组、指标口径、生成时间。
- **兼容性处理**：排名用 `RANK` 而非 `RANK.EQ`，分组最值用 `_xlfn.MAXIFS`，跨表用 `INDEX/MATCH` 而非 `XLOOKUP`。

## 目录结构

```
.
├── SKILL.md                          # 工作流：分流 → 挖掘 → 落地 → 自检 → 交付
├── references/
│   ├── single-table-prompt.md        # 单表版提示词（原文）
│   ├── multi-table-prompt.md         # 多表版提示词（原文）
│   └── formula-recipes.md            # spec 结构、指标类型、公式速查、常见坑
└── scripts/
    ├── inspect_workbook.py           # 读表体检：字段/类型/空值/唯一键/合计行
    ├── add_stats_sheet.py            # 按 spec 生成公式版统计 sheet（含 --check 自检）
    └── recalc_values.sh              # 可选：用 LibreOffice 缓存公式结果，方便预览
```

## 安装

把整个目录放进 Agent 的 user skills 目录即可，例如：

```bash
git clone https://github.com/1162711644/lengmeng-stats-mining.git \
  ~/.doubao/agent_mode/workspace/.user_skills/冷檬数据统计需求挖掘
```

依赖：

```bash
pip install openpyxl
# 可选，仅 recalc_values.sh 需要
# LibreOffice (soffice)
```

## 怎么用

对它说：

> 用冷檬数据统计需求挖掘看看这张表能算什么

它会先做表体检，然后给你四块：

```
读表摘要：这张《销售明细》表有 7 个字段（城市、门店、日期、销售额、数量、订单号、成本），是一份按订单记录的销售流水，共 20 行。

值得统计的业务数据（挑一类，也可以自己说）：
1. 哪个城市/门店卖得多——销售额合计、订单数、平均客单价
2. 卖得好不好赚——销售额成本比、毛利合计
3. 各城市占多大比重——销售额占比、排名
4. 一段时间卖得怎么样——按月/按周的销售额、订单数

你现在最想看哪一类？心里另有想算的也可以直接说。

【做统计】指令（够用版）：
统计：在【销售明细】表里按【城市】分组，算每个城市的【销售额合计】、【订单数】……

用法：把这句整段复制发给我，我直接读表出结果，你不用碰公式。
```

你把指令发回来，它就会输出 `原表_统计结果.xlsx`：原表原样保留，新增 `按城市统计` 等 sheet，全部是公式。

## 脚本单独使用

```bash
# 1. 读表体检
python3 scripts/inspect_workbook.py 数据表.xlsx

# 2. 写 spec.json，生成结果
python3 scripts/add_stats_sheet.py spec.json

# 3. 自检（公式数量 / sheet / fullCalcOnLoad）
python3 scripts/add_stats_sheet.py --check 原表_统计结果.xlsx

# 4. 可选：把公式结果缓存进文件，手机/预览也能看到数字
bash scripts/recalc_values.sh 原表_统计结果.xlsx
```

spec 四种版式（详见 `references/formula-recipes.md`）：

| layout | 用途 | 典型指标 |
| --- | --- | --- |
| `group` | 按维度分组汇总 | 求和、计数、平均、最值、占比、排名、比值 |
| `total` | 整表一个总数 | 求和、平均、最值、行数 |
| `join` | 两表按关联键对齐 | 搬字段 + 行级比值/差值 |
| `combine` | 多表同字段纵向合计 | 合计、篇数、加权平均 |

```json
{
  "source": {"file": "/data/销售明细.xlsx", "sheet": "明细", "header_row": 3, "data_range": "A4:G23"},
  "output": {"file": "/data/销售明细_统计结果.xlsx"},
  "sheets": [
    {
      "name": "按城市统计", "layout": "group", "group_by": ["城市"],
      "metrics": [
        {"label": "销售额合计", "type": "sum", "field": "销售额"},
        {"label": "订单数", "type": "count"},
        {"label": "销售额占比", "type": "share_of", "base": "销售额合计"},
        {"label": "销售额排名", "type": "rank_of", "base": "销售额合计"}
      ]
    }
  ]
}
```

## 已验证

- 用带标题行、"合计"尾行的销售表 + 员工工资表，跑通 `group / total / join / combine` 四种版式；
- LibreOffice 重算后与手工汇总逐项一致（求和、计数、平均、占比、排名、比值、跨表匹配）；
- `fullCalcOnLoad=True` 已设置，Excel/WPS 打开即自动计算。

## 已知边界

- 只吃 `.xlsx / .xlsm`，CSV 请先转 xlsx（体检脚本支持 CSV 临时转换）。
- openpyxl 另存会丢图表 / 透视表 / 图片，遇到会先提醒并给出"独立结果工作簿"选项。
- 只做统计，不出图、不做归因分析；需要看板或报告请配合仪表盘 / 分析报告类 Skill。

## License

MIT
