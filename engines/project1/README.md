# 项目1 · 店小秘加工

## 职责

把原始 Order Excel 交给原样接入的 `vendor/order-convert`，导出：

- 店小秘（dianxiaomi）
- 店管家01（dian_01）
- 店管家02（dian_02）

## 禁止修改（冻结）

以下文件 **禁止** 由 ERP / 封装层改写：

- `vendor/order-convert/src/order_convert/transform.py`
- `vendor/order-convert/src/order_convert/pipeline.py`
- `vendor/order-convert/src/order_convert/resources/rules.yaml`
- `vendor/order-convert/config/rules.yaml`

指纹清单：`VENDOR_MANIFEST.sha256`。改动会让 `test_vendor_package_untouched` 失败。

## 封装层

包路径：`engines/project1/project1_engine/`

```python
import sys
from pathlib import Path
sys.path.insert(0, str(Path("engines/project1").resolve()))
from project1_engine import status, convert, convert_files

status()                       # 就绪 / 版本 / rules 指纹
convert(input_xlsx, out_dir)   # 单文件 → dict（路径为绝对路径）
convert_files([..], out_dir)   # 多文件
```

Backend 薄封装：`backend/app/services/project1_runner.py`（把输出改成相对 `data/`）。

## HTTP 接口（与打单2约定）

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| POST | `/api/project1/upload` | 多文件上传，返回 `batch_id` |
| POST | `/api/project1/run` | body `{ "batch_id": N }` 触发加工 |
| GET | `/api/project1/result/{batch_id}` | 三表下载路径 + 预览摘要 |
| GET | `/api/project1/download?batch_id=&path=` | 下载导出文件 |
| GET | `/api/project1/status` | 引擎状态 |

兼容：`POST /api/batches`（project=project1）与 `GET /api/engines/project1/status` 仍可用。

## 页面流程

`/project1`（店小秘加工）：选文件 →「开始加工」→ 展示状态/统计/告警 → 下载三表。不离开本页。
