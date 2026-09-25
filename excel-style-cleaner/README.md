# Excel Style Cleaner

通用 Excel 样式清洗工具 — 把任何 `.csv / .xls / .xlsx / .xlsm` 文件清洗成"可直接展示"的看板 / 汇总报表。

## 这是什么

针对运营 / 数据分析场景下的**看板 / 月度业绩报表 / 经营分析报表**等汇总类报表：

- 自动识别多层表头、智能合并表头单元格
- 统一字体（微软雅黑 9 号）、边框、居中对齐
- 达成率列 → 黄色 Data Bar
- 环比 / 同比列 → 红色 Data Bar + 上中下箭头
- 金额列 → 万级格式（`0.0万`）+ 负数红
- 表头日期 → `YY年MM月` 格式
- 清理 AI 插件生成的浮层对象 / docProps 水印

## 适用场景 vs 不适用场景

**适合本 skill**：销售看板、月度业绩看板、含合计 / 汇总 / 达成率 / 环比 / 同比的报表，无唯一性主键。

**不适合**（请改用 `excel-data-cleaner`）：订单 / 会员 / 商品 / 物流等**明细表**，含唯一性字段主键。

## 快速开始

```bash
# openpyxl 引擎（默认）—— 清洗旧表
python scripts/excel_style_cleaner.py input.xlsx -o output.xlsx

# xlsxwriter 引擎 —— 纯写新表（BI 导出场景，性能更好）
python scripts/excel_style_cleaner.py input.xlsx -o output.xlsx --xwriter
```

## 项目结构

```
excel-style-cleaner/
├── SKILL.md                      # skill 入口文档（详细架构、参数、决策记录）
├── README.md                     # 本文件（对外项目介绍）
├── images/
│   └── wechat-qrcode.jpg         # 公众号二维码（见下方）
├── scripts/
│   ├── excel_style_cleaner.py    # 主清洗脚本（openpyxl 引擎）
│   └── excel_writer.py           # 纯写新表引擎（xlsxwriter）
├── config/
│   └── keywords.yaml             # 关键字 / 颜色 / 阈值配置
└── references/
    ├── _test_v317.py             # 端到端测试（v3.17 自检 7 项）
    └── smoke_test.md             # 端到端测试（基础 14 项）
```

完整架构、参数、决策记录请阅读 [SKILL.md](./SKILL.md)。

## 微信公众号

在使用本 skill 过程中遇到 bug 或有疑问，欢迎通过下方微信公众号留言：

![微信公众号二维码](./images/wechat-qrcode.jpg)

> 扫码关注，留言会收到回复。

## 版本

当前版本：详见 `SKILL.md` 与 git commit 历史。

## License

MIT
