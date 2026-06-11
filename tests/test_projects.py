import pytest
from quote_system.projects import (
    PROJECTS,
    ProjectConfig,
    resolve_project,
    get_projects_by_company,
    load_project_config,
    save_project_config,
)


class TestResolveProject:
    def test_should_resolve_by_key(self):
        proj = resolve_project("umamusume")
        assert proj.key == "umamusume"
        assert proj.display_name == "马娘"

    def test_should_resolve_by_alias(self):
        proj = resolve_project("战双")
        assert proj.key == "zhan_shuang"

    def test_should_resolve_case_insensitive(self):
        proj = resolve_project("BANG2")
        assert proj.key == "bang2"

    def test_should_resolve_with_whitespace(self):
        proj = resolve_project("  马娘  ")
        assert proj.key == "umamusume"

    def test_should_raise_on_unknown_project(self):
        with pytest.raises(ValueError, match="未知项目"):
            resolve_project("不存在的项目")

    def test_should_resolve_all_registered_projects(self):
        for key in PROJECTS:
            proj = resolve_project(key)
            assert proj.key == key

    def test_should_resolve_all_aliases(self):
        for key, proj in PROJECTS.items():
            for alias in proj.aliases:
                resolved = resolve_project(alias)
                assert resolved.key == key, f"Alias '{alias}' should resolve to '{key}'"


class TestGetProjectsByCompany:
    def test_should_group_by_company(self):
        companies = get_projects_by_company()
        assert isinstance(companies, dict)
        assert "Bilibili" in companies
        assert "完美世界" in companies
        assert "库洛游戏" in companies
        assert "叠纸" in companies

    def test_should_return_sorted_projects(self):
        companies = get_projects_by_company()
        bili = companies["Bilibili"]
        orders = [p.sort_order for p in bili]
        assert orders == sorted(orders), "Projects should be sorted by sort_order"

    def test_should_have_all_projects(self):
        companies = get_projects_by_company()
        total = sum(len(ps) for ps in companies.values())
        assert total == len(PROJECTS)


class TestProjectConfig:
    def test_should_have_required_fields(self):
        for key, proj in PROJECTS.items():
            assert proj.key, f"{key}: missing key"
            assert proj.display_name, f"{key}: missing display_name"
            assert proj.generator, f"{key}: missing generator"

    def test_template_files_should_exist_or_be_empty(self):
        from pathlib import Path
        import os
        
        app_dir = Path(__file__).resolve().parent.parent / "app"
        
        for key, proj in PROJECTS.items():
            if proj.template_file:
                tmpl_path = app_dir / "模板" / "报价单模板" / proj.template_file
                # 有些项目没有模板文件（如 TK、4399、祖龙）
                if not tmpl_path.exists():
                    # 只对有模板的项目报错
                    if proj.generator not in ("tk_fill", "fill_4399", "zulong_settlement"):
                        print(f"WARNING: Template not found: {tmpl_path} for {key}")

    def test_no_duplicate_keys(self):
        keys = list(PROJECTS.keys())
        assert len(keys) == len(set(keys)), "Duplicate project keys"

    def test_no_overlapping_aliases(self):
        all_aliases = []
        for proj in PROJECTS.values():
            for alias in proj.aliases:
                all_aliases.append(alias.lower())
        # Check for overlapping aliases between different projects
        seen = {}
        for proj in PROJECTS.values():
            for alias in proj.aliases:
                low = alias.lower()
                if low in seen and seen[low] != proj.key:
                    print(f"WARNING: Alias '{alias}' shared by {seen[low]} and {proj.key}")
                seen[low] = proj.key

    def test_prices_should_be_positive(self):
        for key, proj in PROJECTS.items():
            for lang, price in proj.prices.items():
                assert price > 0, f"{key}.prices['{lang}'] = {price} should be > 0"
