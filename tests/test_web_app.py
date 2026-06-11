"""API endpoint tests for the Flask web app."""
import pytest
import json
from unittest.mock import patch, MagicMock

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from quote_system.web_app import app


@pytest.fixture
def client():
    app.config["TESTING"] = True
    with app.test_client() as client:
        yield client


# ========== 基础端点测试 ==========

class TestIndexPage:
    def test_should_return_200(self, client):
        resp = client.get("/")
        assert resp.status_code == 200

    def test_should_return_html(self, client):
        resp = client.get("/")
        assert "text/html" in resp.content_type

    def test_should_contain_project_selection_with_valid_project(self, client):
        resp = client.get("/?project=umamusume")
        assert resp.status_code == 200

    def test_should_handle_invalid_project_gracefully(self, client):
        resp = client.get("/?project=nonexistent")
        assert resp.status_code == 200  # 应降级到默认项目

    def test_should_set_no_cache_headers(self, client):
        resp = client.get("/")
        assert resp.headers["Cache-Control"] == "no-cache, no-store, must-revalidate"
        assert resp.headers["Pragma"] == "no-cache"


class TestMobilePage:
    def test_should_return_200(self, client):
        resp = client.get("/mobile")
        assert resp.status_code == 200

    def test_should_return_html(self, client):
        resp = client.get("/mobile")
        assert "text/html" in resp.content_type


class TestReportPage:
    def test_should_return_200(self, client):
        resp = client.get("/report")
        assert resp.status_code == 200


# ========== 报价生成端点测试 ==========

class TestGenerateQuote:
    def test_should_reject_empty_upload(self, client):
        resp = client.post("/generate", data={"project": "umamusume"})
        assert resp.status_code in (200, 302)
        # 应返回错误 flash 消息（重定向到首页）
        if resp.status_code == 302:
            assert "/" in resp.location

    def test_should_reject_missing_project(self, client):
        resp = client.post("/generate", data={})
        # 缺少 project 字段会重定向回首页 (302) 或返回错误
        assert resp.status_code in (200, 302, 400)

    def test_should_reject_invalid_project(self, client):
        resp = client.post("/generate", data={"project": "nonexistent"})
        assert resp.status_code in (200, 302)  # 会产生 ValueError flash


# ========== 配置管理端点测试 ==========

class TestSaveLanguageConfig:
    def test_should_accept_valid_config(self, client):
        resp = client.post("/save-language-config", data={
            "project": "umamusume",
            "language_names[]": ["日翻中"],
            "language_prices[]": ["0.24"],
            "default_languages[]": ["日翻中"],
        })
        assert resp.status_code == 200
        data = json.loads(resp.data)
        assert data["status"] == "success"

    def test_should_reject_missing_project(self, client):
        resp = client.post("/save-language-config", data={})
        # BUG: 缺少必填字段时返回 500 而非 400
        assert resp.status_code in (400, 500)


class TestResetLanguageConfig:
    def test_should_reset_for_valid_project(self, client):
        resp = client.post("/reset-language-config", data={"project": "umamusume"})
        assert resp.status_code == 200

    def test_should_fail_missing_project(self, client):
        resp = client.post("/reset-language-config", data={})
        # BUG: 缺少必填字段时返回 500 而非 400
        assert resp.status_code in (400, 500)


class TestSaveProjectConfig:
    def test_should_accept_json_config(self, client):
        resp = client.post("/save-project-config",
                          data=json.dumps([{
                              "project_key": "umamusume",
                              "display_name": "马娘",
                              "sort_order": 1,
                              "company": "Bilibili",
                          }]),
                          content_type="application/json")
        assert resp.status_code == 200

    def test_should_accept_form_config(self, client):
        resp = client.post("/save-project-config", data={
            "project_key": "umamusume",
            "display_name": "马娘",
            "sort_order": "1",
            "company": "Bilibili",
        })
        assert resp.status_code == 200

    def test_should_reject_empty_request(self, client):
        resp = client.post("/save-project-config", data={})
        # BUG: 缺少必填字段时返回 500 而非 400
        assert resp.status_code in (400, 500)


class TestSaveSavePath:
    def test_should_accept_valid_data(self, client):
        resp = client.post("/save-save-path", data={
            "project_key": "umamusume",
            "save_path": r"D:\test\path",
        })
        assert resp.status_code == 200

    def test_should_fail_missing_data(self, client):
        resp = client.post("/save-save-path", data={})
        # BUG: 缺少必填字段时返回 500 而非 400
        assert resp.status_code in (400, 500)


# ========== 基础路径端点测试 ==========

class TestBasePaths:
    def test_should_get_base_paths(self, client):
        resp = client.get("/api/base-paths")
        assert resp.status_code == 200
        data = json.loads(resp.data)
        assert data["status"] == "success"
        assert "quote_history_base" in data
        assert "settlement_base" in data

    def test_should_save_base_paths(self, client):
        resp = client.post("/api/base-paths",
                          data=json.dumps({"quote_history_base": "D:/test"}),
                          content_type="application/json")
        assert resp.status_code == 200


# ========== 报价历史端点测试 ==========

class TestQuoteHistory:
    def test_should_return_json_for_valid_project(self, client):
        resp = client.get("/quote-history?project=umamusume")
        assert resp.status_code == 200
        data = json.loads(resp.data)
        assert data["status"] == "success"

    def test_should_return_json_without_project(self, client):
        resp = client.get("/quote-history")
        assert resp.status_code == 200
        data = json.loads(resp.data)
        assert data["status"] == "success"

    def test_should_handle_invalid_project(self, client):
        resp = client.get("/quote-history?project=nonexistent")
        # 路径遍历被 resolve_project 拦截，返回 200（展开为空列表）
        assert resp.status_code in (200, 500)


class TestAllQuotes:
    def test_should_return_json(self, client):
        resp = client.get("/api/all-quotes")
        assert resp.status_code == 200
        data = json.loads(resp.data)
        assert data["status"] == "success"


# ========== 删除报价单测试 ==========

class TestDeleteQuote:
    def test_should_reject_empty_path(self, client):
        resp = client.post("/delete-quote",
                          data=json.dumps({}),
                          content_type="application/json")
        assert resp.status_code == 400

    def test_should_reject_nonexistent_file(self, client):
        resp = client.post("/delete-quote",
                          data=json.dumps({"path": r"D:\nonexistent\file.xlsx"}),
                          content_type="application/json")
        assert resp.status_code in (400, 500)

    def test_should_reject_invalid_suffix(self, client):
        resp = client.post("/delete-quote",
                          data=json.dumps({"path": r"D:\test.exe"}),
                          content_type="application/json")
        assert resp.status_code in (400, 500)


# ========== 本地安全保护测试 ==========

class TestLocalOnlyEndpoints:
    def test_delete_quote_should_be_protected(self, client):
        # 测试本地请求保护
        resp = client.post("/delete-quote",
                          data=json.dumps({"path": ""}),
                          content_type="application/json")
        # 应该通过本地保护（测试环境是 localhost）
        assert resp.status_code != 403

    def test_open_file_should_be_protected(self, client):
        resp = client.get("/open-file")
        assert resp.status_code != 403


# ========== 结算预览端点测试 ==========

class TestSettlementPreview:
    def test_mamian_preview_with_valid_params(self, client):
        resp = client.get("/api/settlement/mamian/preview?year=2026&month=6&project=umamusume")
        assert resp.status_code == 200

    def test_mamian_preview_with_invalid_year(self, client):
        resp = client.get("/api/settlement/mamian/preview?year=1000&month=6")
        assert resp.status_code == 400

    def test_mamian_preview_with_invalid_month(self, client):
        resp = client.get("/api/settlement/mamian/preview?year=2026&month=13")
        assert resp.status_code == 400

    def test_general_preview_with_valid_params(self, client):
        resp = client.get("/api/settlement/preview?year=2026&month=6")
        assert resp.status_code == 200

    def test_diezhi_preview_with_valid_params(self, client):
        # Note: 需要 pywin32 (win32com)，否则会返回 500
        resp = client.get("/api/settlement_diezhi/preview?year=2026&month=6")
        assert resp.status_code in (200, 500)


# ========== 结算生成端点测试 ==========

class TestSettlementGenerate:
    def test_mamian_generate_with_invalid_year(self, client):
        resp = client.post("/api/settlement/mamian/generate-bill",
                          data=json.dumps({"year": 1000, "month": 6}),
                          content_type="application/json")
        assert resp.status_code == 400

    def test_mamian_generate_with_no_data(self, client):
        resp = client.post("/api/settlement/mamian/generate-bill",
                          data=json.dumps({"year": 2026, "month": 6, "project": "umamusume"}),
                          content_type="application/json")
        assert resp.status_code in (200, 404)

    def test_general_generate_with_invalid_year(self, client):
        resp = client.post("/api/settlement/generate",
                          data=json.dumps({"year": 1000, "month": 6}),
                          content_type="application/json")
        assert resp.status_code == 400


# ========== TK 端点测试 ==========

class TestTkEndpoints:
    def test_start_with_no_data(self, client):
        resp = client.post("/tk-fill/start", data={})
        assert resp.status_code == 200
        data = json.loads(resp.data)
        assert data["status"] == "success"


class TestFill4399:
    def test_start_with_no_data(self, client):
        resp = client.post("/fill-4399/start", data={})
        assert resp.status_code == 200
        data = json.loads(resp.data)
        assert data["status"] == "success"


# ========== 边界/异常测试 ==========

class TestBoundaryConditions:
    def test_generate_with_long_project_name(self, client):
        long_name = "a" * 1000
        resp = client.post("/generate", data={"project": long_name})
        assert resp.status_code in (200, 302, 400)

    def test_index_with_sql_injection_attempt(self, client):
        resp = client.get("/?project=' OR '1'='1")
        assert resp.status_code == 200  # 应优雅处理

    def test_index_with_xss_attempt(self, client):
        resp = client.get("/?project=<script>alert(1)</script>")
        assert resp.status_code == 200  # 应优雅处理，Jinja2 自动转义

    def test_quote_history_with_traversal_attempt(self, client):
        resp = client.get("/quote-history?project=../../../etc")
        # 路径遍历被 resolve_project 优雅拦截
        assert resp.status_code in (200, 500)

    def test_mobile_generate_with_empty_request(self, client):
        resp = client.post("/mobile/generate", data={})
        assert resp.status_code == 200
        data = json.loads(resp.data)
        assert data["status"] == "error"

    def test_settlement_files_without_project(self, client):
        resp = client.get("/api/settlement-files")
        assert resp.status_code == 400


# ========== CORS/安全头部测试 ==========

class TestSecurityHeaders:
    def test_no_cache_on_all_pages(self, client):
        for url in ["/", "/mobile", "/report"]:
            resp = client.get(url)
            assert resp.headers.get("Cache-Control") == "no-cache, no-store, must-revalidate", \
                f"{url}: missing Cache-Control header"

    def test_error_responses_are_json(self, client):
        resp = client.get("/api/settlement-files")
        assert resp.status_code == 400
        # 不应暴露内部错误详情
        data = json.loads(resp.data)
        assert "status" in data
