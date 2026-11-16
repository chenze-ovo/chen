"""项目1专用 upload → run → result 接口。"""
import io

from openpyxl import Workbook

from tests.conftest import login


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


def test_project1_upload_run_result_download(client, admin_headers):
    content = _xlsx([["P1-API-1", "shop", "SIMPLE", "云途", "TRK1", "", 1, 10, "Ann", "US", "addr", "1", ""]])
    uploaded = client.post(
        "/api/project1/upload",
        headers=admin_headers,
        data={"file_kind": "raw_order"},
        files=[
            (
                "files",
                (
                    "raw.xlsx",
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ),
            )
        ],
    )
    assert uploaded.status_code == 201, uploaded.text
    body = uploaded.json()
    assert body["status"] == "created"
    batch_id = body["batch_id"]

    ran = client.post(
        "/api/project1/run",
        headers=admin_headers,
        json={"batch_id": batch_id},
    )
    assert ran.status_code == 200, ran.text
    assert ran.json()["batch_id"] == batch_id
    assert ran.json()["status"] in {"converted", "needs_review", "blocked", "failed"}
    assert isinstance(ran.json()["downloads"], list)

    result = client.get(f"/api/project1/result/{batch_id}", headers=admin_headers)
    assert result.status_code == 200, result.text
    detail = result.json()
    assert detail["batch_id"] == batch_id
    assert detail["files"]
    assert detail["files"][0]["export_status"] in {"ok", "needs_review", "blocked", "error"}
    assert "stats" in detail["files"][0]

    if detail["downloads"]:
        path = detail["downloads"][0]["path"]
        fetched = client.get(
            "/api/project1/download",
            headers=admin_headers,
            params={"batch_id": batch_id, "path": path},
        )
        assert fetched.status_code == 200
        assert len(fetched.content) > 50


def test_project1_run_requires_ops(client):
    reviewer = login(client, "reviewer", "reviewer123")
    denied = client.post("/api/project1/run", headers=reviewer, json={"batch_id": 1})
    assert denied.status_code == 403


def test_project1_status_endpoint(client, admin_headers):
    response = client.get("/api/project1/status", headers=admin_headers)
    assert response.status_code == 200
    body = response.json()
    assert body["project"] == "project1"
    assert body["ready"] is True



def _engine():
    import sys
    from pathlib import Path

    root = Path(__file__).resolve().parents[2] / "engines" / "project1"
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    from project1_engine.engine import _issue_brief, collect_visible_issues, explain_issue

    return _issue_brief, collect_visible_issues, explain_issue


def test_issue_brief_keeps_sku_and_suffix():
    brief, _, explain = _engine()
    rows = brief(
        [
            {
                "type": "UNKNOWN_SKU_SUFFIX",
                "severity": "warning",
                "订单号": "A-1",
                "action": "已去掉未知SKU后缀，仅输出基础SKU；需确认后缀白名单，禁止正式导入",
                "SKU": "M0021-CQCD",
                "suffix": "CQCD",
            }
        ]
    )
    assert len(rows) == 1
    assert rows[0]["SKU"] == "M0021-CQCD"
    assert rows[0]["suffix"] == "CQCD"
    assert rows[0]["订单号"] == "A-1"
    assert rows[0]["type"] == "UNKNOWN_SKU_SUFFIX"
    wording = explain(rows[0])
    assert "CQCD" in wording["problem"]
    assert "M0021-CQCD" in wording["problem"]
    assert "M0021" in wording["problem"]
    assert "记下这个后缀" in wording["how_to"]
    assert wording.get("action") == {"type": "record_suffix", "suffix": "CQCD"}
    assert "UNKNOWN_SKU_SUFFIX" not in wording["problem"]
    assert "UNKNOWN_SKU_SUFFIX" not in wording["how_to"]


def test_issue_brief_does_not_silently_truncate():
    brief, collect, _ = _engine()
    items = [
        {
            "type": "UNKNOWN_SKU_SUFFIX",
            "severity": "warning",
            "订单号": f"ORD-{i}",
            "SKU": f"M0021-CQCD",
            "suffix": "CQCD",
            "action": "已去掉未知SKU后缀",
        }
        for i in range(40)
    ]
    rows = brief(items)
    assert len(rows) == 40
    assert rows[0]["suffix"] == "CQCD"
    assert rows[39]["订单号"] == "ORD-39"
    visible = collect([{"file": "a.xlsx", "errors": [], "warnings": rows}])
    assert len(visible) == 40


def test_info_issues_are_not_listed():
    _, collect, _ = _engine()
    visible = collect(
        [
            {
                "file": "a.xlsx",
                "errors": [],
                "warnings": [
                    {
                        "type": "SKU_FROM_NOTE",
                        "severity": "info",
                        "订单号": "A-1",
                        "message": "从备注取了货号",
                    },
                    {
                        "type": "MULTI_SKU",
                        "severity": "warning",
                        "订单号": "A-1",
                        "SKU": "M1\nM2",
                        "sku_lines": ["M1", "M2"],
                    },
                    {
                        "type": "MULTI_SKU",
                        "severity": "warning",
                        "订单号": "A-1",
                        "SKU": "M1\nM2",
                        "sku_lines": ["M1", "M2"],
                    },
                ],
            }
        ]
    )
    assert len(visible) == 1
    assert "多个货号" in visible[0]["problem"]
    assert "SKU_FROM_NOTE" not in visible[0]["problem"]


def test_result_api_returns_suffix_details_and_plain_issues(client, admin_headers):
    content = _xlsx(
        [["P1-SUF-1", "Shirley", "M0021-CQCD", "云途", "TRK-SUF-1", "", 1, 10, "Ann", "US", "addr", "1", ""]]
    )
    uploaded = client.post(
        "/api/project1/upload",
        headers=admin_headers,
        data={"file_kind": "raw_order"},
        files=[
            (
                "files",
                (
                    "raw-suffix.xlsx",
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ),
            )
        ],
    )
    assert uploaded.status_code == 201, uploaded.text
    batch_id = uploaded.json()["batch_id"]
    ran = client.post("/api/project1/run", headers=admin_headers, json={"batch_id": batch_id})
    assert ran.status_code == 200, ran.text
    result = client.get(f"/api/project1/result/{batch_id}", headers=admin_headers)
    assert result.status_code == 200, result.text
    detail = result.json()
    assert detail["exception_count"] >= 1
    assert detail["issues"]
    issue = detail["issues"][0]
    assert "CQCD" in issue["problem"]
    assert "M0021-CQCD" in issue["problem"]
    assert "记下这个后缀" in issue["how_to"]
    assert issue.get("action", {}).get("type") == "record_suffix"
    assert "UNKNOWN_SKU_SUFFIX" not in issue["problem"]
    warnings = detail["files"][0]["warnings"]
    suffix_rows = [item for item in warnings if item.get("type") == "UNKNOWN_SKU_SUFFIX"]
    assert suffix_rows
    assert suffix_rows[0]["SKU"] == "M0021-CQCD"
    assert suffix_rows[0]["suffix"] == "CQCD"
    notes = [item for item in detail["downloads"] if item.get("kind") == "notes" or item.get("label") == "异常说明"]
    assert notes
    fetched = client.get(
        "/api/project1/download",
        headers=admin_headers,
        params={"batch_id": batch_id, "path": notes[0]["path"]},
    )
    assert fetched.status_code == 200, fetched.text
    body = fetched.content.decode("utf-8")
    assert "CQCD" in body
    assert "M0021-CQCD" in body

    listed = client.get("/api/batches?page_size=5", headers=admin_headers)
    assert listed.status_code == 200
    rows = listed.json()["items"]
    this = next(item for item in rows if item["id"] == batch_id)
    assert this["exception_count"] >= 1


def test_result_hydrates_suffix_from_anomalies_json(client, admin_headers):
    content = _xlsx(
        [["P1-SUF-2", "Shirley", "M0021-CQCD", "云途", "TRK-SUF-2", "", 1, 10, "Ann", "US", "addr", "1", ""]]
    )
    uploaded = client.post(
        "/api/project1/upload",
        headers=admin_headers,
        data={"file_kind": "raw_order"},
        files=[
            (
                "files",
                (
                    "raw-hydrate.xlsx",
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ),
            )
        ],
    )
    assert uploaded.status_code == 201, uploaded.text
    batch_id = uploaded.json()["batch_id"]
    ran = client.post("/api/project1/run", headers=admin_headers, json={"batch_id": batch_id})
    assert ran.status_code == 200, ran.text

    from app.database import SessionLocal
    from app.models import ImportBatch
    from sqlalchemy.orm.attributes import flag_modified

    db = SessionLocal()
    try:
        batch = db.get(ImportBatch, batch_id)
        assert batch is not None
        summary = dict(batch.engine_summary or {})
        files = []
        for item in summary.get("files") or []:
            slim = dict(item)
            slim["warnings"] = [
                {
                    "type": w.get("type"),
                    "severity": w.get("severity"),
                    "订单号": w.get("订单号"),
                    "message": w.get("message") or w.get("action") or "",
                }
                for w in (item.get("warnings") or [])
            ]
            slim["errors"] = [
                {
                    "type": e.get("type"),
                    "severity": e.get("severity"),
                    "订单号": e.get("订单号"),
                    "message": e.get("message") or e.get("action") or "",
                }
                for e in (item.get("errors") or [])
            ]
            slim["anomalies"] = [
                {
                    "type": a.get("type"),
                    "severity": a.get("severity"),
                    "订单号": a.get("订单号"),
                    "message": a.get("message") or a.get("action") or "",
                }
                for a in (item.get("anomalies") or [])
            ]
            files.append(slim)
        summary["files"] = files
        batch.engine_summary = summary
        flag_modified(batch, "engine_summary")
        db.add(batch)
        db.commit()
    finally:
        db.close()

    result = client.get(f"/api/project1/result/{batch_id}", headers=admin_headers)
    assert result.status_code == 200, result.text
    detail = result.json()
    warnings = detail["files"][0]["warnings"]
    suffix_rows = [item for item in warnings if item.get("type") == "UNKNOWN_SKU_SUFFIX"]
    assert suffix_rows, warnings
    assert suffix_rows[0].get("SKU") == "M0021-CQCD"
    assert suffix_rows[0].get("suffix") == "CQCD"
    assert any("CQCD" in (issue.get("problem") or "") for issue in detail["issues"])
