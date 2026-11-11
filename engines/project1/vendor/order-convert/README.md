# 订单 Excel 自动转换（核心库）

从一份原始 Order Excel 生成店小秘 / 店管家01 / 店管家02，并输出统计、异常与行级映射。  
**无 Windows GUI/EXE**；规则见 `config/rules.yaml`（包内亦有默认副本）。

## 从零安装

```bash
cd order-convert   # 或解压交付包后的目录
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

## 转换（默认使用包内空白模板，不必再选 dian_01/02 样板）

```bash
PYTHONPATH=src .venv/bin/python -m order_convert -i raw_order.xlsx -o out
# 只导出一份： --only dx|d01|d02
# 可选覆盖模板： --template-d01 /path/to.xls
```

退出码：`0` 全部确认可正式导出；`1` 严重错误（blocked）；`2` 待复核（needs_review）。

## 导出状态三分法（`stats.export_status`）

| 状态 | 含义 | 写出位置 |
|---|---|---|
| **ok** | 全部规则确认 | `out/dianxiaomi.xlsx`、`out/dian_01.xls`、`out/dian_02.xls` |
| **needs_review** | 未知 SKU 后缀、多 SKU（换行）等未确认 | **仅** `out/review/*_NEEDS_REVIEW.*`；正式可导入文件名**不会**出现在 `out/` 根目录 |
| **blocked** | 未知店铺、缺列、重复单号等严重错误 | 不写正式 01/02；店小秘可写预览到 `out/`；可选 `--allow-review-export` 写 review |

**SKU 后缀**：输出一律去掉最后一段 `-后缀`（白名单 `LYF`/`XYTU`、未知后缀、以及多 SKU 各行均同）；日本标签（`-LYFJP` + 拣货备注含关键字）店小秘仍为 `基础SKU + 备注`，02 用基础 SKU。未知后缀记 `UNKNOWN_SKU_SUFFIX`，**仍 needs_review**，禁止正式导入。

多 SKU：按换行分别去后缀并写入 anomalies；输出用换行拼接去后缀后的各行；**不**擅自拆销售明细、**不**臆造各 SKU 数量；订单标 `MULTI_SKU` 待复核。

## 测试

```bash
PYTHONPATH=src .venv/bin/python -m pytest tests/ -v --tb=short | tee tests_result.txt
```

含 41 单回归（须与人工样板 0 差异）+ 边界/合成夹具用例 + 字体格式断言。

## 内置模板

| 资源 | 路径 |
|---|---|
| 店管家01 空白 | `src/order_convert/templates/dian_01_blank.xls` |
| 店管家02 空白 | `src/order_convert/templates/dian_02_blank.xls` |
| 店小秘 空白 | `src/order_convert/templates/dianxiaomi_blank.xlsx` |
| 默认规则 | `src/order_convert/resources/rules.yaml` |

**如何更新模板**：用新的人工样板覆盖项目根目录 `dian_01.xls` / `dian_02.xls` / `dianxiaomi.xlsx` 后执行：

```bash
PYTHONPATH=src .venv/bin/python scripts/extract_blank_templates.py
```

模板只保留表头 + 列宽（店管家不再残留被 `write("")` 冲成 Arial 的空数据行）。

## 格式保留方案

- **店小秘**：复制空白 xlsx → 数据行强制 **宋体 11**、居中；亚马逊整行红字 ARGB **`FFFF0000`**；非亚马逊黑字 `FF000000`。
- **店管家 01/02**：`xlrd` + `xlutils.copy` 空白模板（保留表头样式与列宽）→ 数据行用 xlwt **显式宋体**样式写入（01 有内容列约 11pt；02 红字列 `colour_index=10` + 宋体 11pt）。
- **xls / BIFF 限制（无法 100% 还原）**：
  - 不能完美克隆全部 Excel 特性（主题色、条件格式、部分边框/行高细节）。
  - xlwt 调色板红 ≠ OOXML 的 `FFFF0000` 语义，但 Excel 中均显示为标准红。
  - 表头「微软雅黑」等来自模板复制；数据行按人工样板重建为宋体，不依赖已清空行的 XF。
  - 长单号按文本写入，避免科学计数；与样板个别单元格「数值型」存储可能不完全同型，显示一致。

## 输出

- 正式（ok）：`dianxiaomi.xlsx` / `dian_01.xls` / `dian_02.xls`
- 待复核（needs_review）：`review/dianxiaomi_NEEDS_REVIEW.xlsx` 等
- 侧车：`stats.json`（含 `export_status`）/ `anomalies.json` / `row_mapping.csv`
