"""中国直发：在异常里记下后缀/店铺后，再转不再提示，并打出店管家表。"""
import io
import zipfile

from openpyxl import Workbook

from app.services.project1_export import build_project1_zip_bytes, is_internal_result_file


def _xlsx(rows):
    book = Workbook()
    sheet = book.active
    sheet.append(
        [
            "订单号",
            "店铺账号",
            "SKU",
            "物流方式",
            "运单号",
            "拣货备注",
            "产品总数",
            "订单金额",
            "收货人姓名",
            "收货人国家",
            "详细地址",
            "收货人电话",
            "收货人手机",
        ]
    )
    for row in rows:
        sheet.append(row)
    buffer = io.BytesIO()
    book.save(buffer)
    return buffer.getvalue()


def _upload_and_run(client, headers, content, name="raw.xlsx"):
    uploaded = client.post(
        "/api/project1/upload",
        headers=headers,
        data={"file_kind": "raw_order"},
        files=[("files", (name, content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))],
    )
    assert uploaded.status_code == 201, uploaded.text
    batch_id = uploaded.json()["batch_id"]
    ran = client.post("/api/project1/run", headers=headers, json={"batch_id": batch_id})
    assert ran.status_code == 200, ran.text
    detail = client.get(f"/api/project1/result/{batch_id}", headers=headers)
    assert detail.status_code == 200, detail.text
    return batch_id, detail.json()


def _issues(detail):
    return detail.get("issues") or []


def test_explain_shop_issue_has_record_action():
    from pathlib import Path
    import sys

    root = Path(__file__).resolve().parents[2] / "engines" / "project1"
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    from project1_engine.engine import explain_issue

    wording = explain_issue(
        {
            "type": "UNKNOWN_SHOP",
            "severity": "error",
            "订单号": "A-1",
            "店铺账号": "香港星辉美客多（自营）",
        }
    )
    assert "香港星辉美客多（自营）" in wording["problem"]
    assert "确认后告诉我们" not in wording["how_to"]
    assert wording["action"] == {"type": "record_shop", "shop": "香港星辉美客多（自营）"}


def test_record_suffix_and_shop_then_dian_01(client, admin_headers):
    content = _xlsx(
        [
            [
                "A-1",
                "香港星辉美客多（自营）",
                "M0021-CQCD",
                "云途",
                "TRK1",
                "",
                1,
                10,
                "Ann",
                "US",
                "addr",
                "1",
                "1",
            ]
        ]
    )
    batch_id, detail = _upload_and_run(client, admin_headers, content)
    types = {item.get("action", {}).get("type") or "" for item in _issues(detail)}
    texts = " ".join(f"{item.get('problem','')} {item.get('how_to','')}" for item in _issues(detail))
    assert "记下这个后缀" in texts or any(item.get("action", {}).get("type") == "record_suffix" for item in _issues(detail))
    assert "record_shop" in types or "美客多" in texts
    downloads = {item.get("kind") for item in detail.get("downloads") or []}
    assert "dx" in downloads

    suffix = client.post(
        "/api/project1/allowlist/suffix",
        headers=admin_headers,
        json={"suffix": "CQCD"},
    )
    assert suffix.status_code == 200, suffix.text
    shop = client.post(
        "/api/project1/allowlist/shop",
        headers=admin_headers,
        json={"shop_account": "香港星辉美客多（自营）", "kind": "other"},
    )
    assert shop.status_code == 200, shop.text

    ran = client.post("/api/project1/run", headers=admin_headers, json={"batch_id": batch_id})
    assert ran.status_code == 200, ran.text
    detail = client.get(f"/api/project1/result/{batch_id}", headers=admin_headers).json()
    texts = " ".join(f"{item.get('problem','')}" for item in _issues(detail))
    assert "CQCD" not in texts
    assert "美客多" not in texts
    kinds = {item.get("kind") for item in detail.get("downloads") or []}
    assert "d01" in kinds
    assert "dx" in kinds


def test_record_shop_as_amazon(client, admin_headers):
    content = _xlsx(
        [
            [
                "B-1",
                "香港星辉美客多（自营）",
                "M0021",
                "云途",
                "TRK2",
                "",
                1,
                10,
                "Ann",
                "US",
                "addr",
                "1",
                "1",
            ]
        ]
    )
    batch_id, _detail = _upload_and_run(client, admin_headers, content, name="amazon.xlsx")
    shop = client.post(
        "/api/project1/allowlist/shop",
        headers=admin_headers,
        json={"shop_account": "香港星辉美客多（自营）", "kind": "amazon"},
    )
    assert shop.status_code == 200, shop.text
    client.post("/api/project1/run", headers=admin_headers, json={"batch_id": batch_id})
    detail = client.get(f"/api/project1/result/{batch_id}", headers=admin_headers).json()
    kinds = {item.get("kind") for item in detail.get("downloads") or []}
    assert "d02" in kinds
    texts = " ".join(item.get("problem", "") for item in _issues(detail))
    assert "美客多" not in texts


def test_zip_uses_chinese_names_and_drops_sidecars(tmp_path):
    stem = tmp_path / "order"
    stem.mkdir()
    (stem / "dianxiaomi.xlsx").write_bytes(b"xlsx")
    (stem / "dian_01.xls").write_bytes(b"xls")
    (stem / "dian_02.xls").write_bytes(b"xls")
    (stem / "异常说明.txt").write_text("说明", encoding="utf-8")
    (stem / "anomalies.json").write_text("[]", encoding="utf-8")
    (stem / "stats.json").write_text("{}", encoding="utf-8")
    (stem / "row_mapping.json").write_text("{}", encoding="utf-8")
    (stem / "row_mapping.csv").write_text("a,b\n", encoding="utf-8")

    assert is_internal_result_file(stem / "anomalies.json")
    payload, filename = build_project1_zip_bytes(tmp_path, batch_id=9)
    assert filename.startswith("中国直发加工结果-批次9-")
    with zipfile.ZipFile(io.BytesIO(payload)) as zf:
        names = set(zf.namelist())
    assert names == {"店小秘.xlsx", "店管家01.xls", "店管家02.xls", "异常说明.txt"}



def test_unknown_suffix_still_downloads_dian(client, admin_headers):
    content = _xlsx(
        [
            [
                "C-1",
                "亚马逊测试店",
                "M0021-CQCD",
                "云途",
                "TRK3",
                "",
                1,
                10,
                "Ann",
                "US",
                "addr",
                "1",
                "1",
            ]
        ]
    )
    _batch_id, detail = _upload_and_run(client, admin_headers, content, name="suffix.xlsx")
    kinds = {item.get("kind") for item in detail.get("downloads") or []}
    assert "dx" in kinds
    assert "d02" in kinds
    assert any((item.get("action") or {}).get("suffix") == "CQCD" for item in _issues(detail))


def test_record_correct_sku_then_no_suffix_issue(client, admin_headers):
    content = _xlsx(
        [
            [
                "D-1",
                "亚马逊测试店",
                "WX12054-HKCQ",
                "云途",
                "TRK4",
                "",
                1,
                10,
                "Ann",
                "US",
                "addr",
                "1",
                "1",
            ]
        ]
    )
    batch_id, detail = _upload_and_run(client, admin_headers, content, name="keep.xlsx")
    assert any((item.get("action") or {}).get("suffix") == "HKCQ" for item in _issues(detail))
    recorded = client.post(
        "/api/project1/allowlist/sku",
        headers=admin_headers,
        json={"original_sku": "WX12054-HKCQ", "corrected_sku": "WX12054-HKCQ"},
    )
    assert recorded.status_code == 200, recorded.text
    client.post("/api/project1/run", headers=admin_headers, json={"batch_id": batch_id})
    detail = client.get(f"/api/project1/result/{batch_id}", headers=admin_headers).json()
    texts = " ".join(item.get("problem", "") for item in _issues(detail))
    assert "HKCQ" not in texts
    kinds = {item.get("kind") for item in detail.get("downloads") or []}
    assert "d02" in kinds
