from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path


@dataclass
class ProjectConfig:
    key: str
    display_name: str
    aliases: tuple[str, ...]
    template_file: str
    history_dir: str
    default_languages: tuple[str, ...]
    prices: dict[str, float]
    generator: str
    file_prefix: str
    company: str = ""
    sort_order: int = 0
    sheet_id: str = ""
    spreadsheet_token: str = ""


PROJECTS: dict[str, ProjectConfig] = {
    "umamusume": ProjectConfig(
        key="umamusume",
        display_name="马娘",
        aliases=("马娘", "代号PD", "pd", "PD", "umamusume"),
        template_file="马娘_报价单模板.xlsx",
        history_dir="Bilibili/马娘",
        default_languages=("日翻中",),
        prices={"中翻日": 0.58, "日翻中": 0.24, "摘字": 0.072},
        generator="punctuation_30",
        file_prefix="报价单_代号PD",
        company="Bilibili",
        sort_order=1,
    ),
    "hbr": ProjectConfig(
        key="hbr",
        display_name="炽焰天穹",
        aliases=("炽焰天穹",),
        template_file="HBR_报价单模板.xlsx",
        history_dir="Bilibili/HBR",
        default_languages=("日翻中",),
        prices={"中翻日": 0.58, "日翻中": 0.24, "摘字": 0.072},
        generator="punctuation_30",
        file_prefix="报价单_炽焰天穹",
        company="Bilibili",
        sort_order=1,
    ),
    "bang2": ProjectConfig(
        key="bang2",
        display_name="BANG2",
        aliases=("BANG2", "bang2", "邦2"),
        template_file="bang2_报价单模板.xlsx",
        history_dir="Bilibili/bang2",
        default_languages=("日翻中", "日翻韩", "日翻英"),
        prices={
            "日翻中": 0.30,
            "日翻韩": 0.81,
            "日翻英": 0.81,
            "日翻繁": 0.45,
            "中翻繁": 0.26,
            "中翻英": 0.68,
            "中翻韩": 0.60,
            "摘字": 0.072,
        },
        generator="bang2",
        file_prefix="报价单_BANG2",
        company="Bilibili",
        sort_order=2,
    ),
    "zhan_shuang": ProjectConfig(
        key="zhan_shuang",
        display_name="战双版更",
        aliases=("战双", "战双版更", "库洛", "zhan_shuang", "zhanshuang"),
        template_file="战双_报价单模板.xlsx",
        history_dir="库洛游戏/战双版更",
        default_languages=("中译韩",),
        prices={"中译韩": 0.64, "中-韩": 0.64},
        generator="zhan_shuang",
        file_prefix="【报价单】战双_中译韩",
        company="库洛游戏",
        sort_order=3,
        sheet_id="WmpiFq",
        spreadsheet_token="Wup0wnUPIiiIr2k8T23c4zASnjd",
    ),
    "zhan_shuang_faxing": ProjectConfig(
        key="zhan_shuang_faxing",
        display_name="战双发行",
        aliases=("战双发行", "库洛发行"),
        template_file="2026年4＆5月发行本地化沟通群需求报价单模板.xlsx",
        history_dir="库洛游戏/战双发行",
        default_languages=(),
        prices={},
        generator="zhan_shuang_feishu",
        file_prefix="",
        company="库洛游戏",
        sort_order=4,
        sheet_id="WmpiFq",
        spreadsheet_token="Wup0wnUPIiiIr2k8T23c4zASnjd",
    ),
    "huanta": ProjectConfig(
        key="huanta",
        display_name="幻塔",
        aliases=("幻塔", "TOF", "huanta", "HuanTa", "完美世界"),
        template_file="幻塔报价_模板.xlsx",
        history_dir="完美世界/幻塔",
        default_languages=("翻译+润色", "润色"),
        prices={"翻译+润色": 0.64, "润色": 0.50},
        generator="huanta",
        file_prefix="报价单_幻塔",
        company="完美世界",
        sort_order=4,
        sheet_id="ZXBogz",
    ),
    "yihuan_nei": ProjectConfig(
        key="yihuan_nei",
        display_name="异环游戏内",
        aliases=("异环游戏内", "异环", "NTE", "Neverness to Everness"),
        template_file="幻塔报价_模板.xlsx",
        history_dir="完美世界/异环游戏内",
        default_languages=("翻译+润色", "润色"),
        prices={"翻译+润色": 0.64, "润色": 0.50},
        generator="huanta",
        file_prefix="报价单_异环游戏内",
        company="完美世界",
        sort_order=5,
        sheet_id="bfba7c",
    ),
    "yihuan_faxing": ProjectConfig(
        key="yihuan_faxing",
        display_name="异环发行",
        aliases=("异环发行", "异环发行", "NTE发行"),
        template_file="幻塔报价_模板.xlsx",
        history_dir="完美世界/异环发行",
        default_languages=("翻译+润色", "润色"),
        prices={"翻译+润色": 0.64, "润色": 0.50, "翻译+润色英日": 0.72},
        generator="huanta",
        file_prefix="报价单_异环发行",
        company="完美世界",
        sort_order=6,
        sheet_id="S5yHmP",
    ),
    "liandishenkong": ProjectConfig(
        key="liandishenkong",
        display_name="恋与深空",
        aliases=("恋与深空", "X3", "x3", "liandishenkong"),
        template_file="批次40-报价单_恋与深空_中译日模板.xlsx",
        history_dir="叠纸/恋与深空",
        default_languages=("中译日",),
        prices={"中译日": 0.52},
        generator="diezhi",
        file_prefix="批次40-报价单_恋与深空_中译日",
        company="叠纸",
        sort_order=8,
        sheet_id="4aca99",
        spreadsheet_token="JtXvsAgthhLA9LtZw2qcswfZnkd",
    ),
    "shining_nikki": ProjectConfig(
        key="shining_nikki",
        display_name="闪暖",
        aliases=("闪暖", "shining_nikki", "SN"),
        template_file="第42批-报价单_闪暖_中译日模板.xlsx",
        history_dir="叠纸/闪暖",
        default_languages=("初翻",),
        prices={"初翻": 0.32, "初翻+审校": 0.64},
        generator="diezhi",
        file_prefix="第42批-报价单_闪暖_中译日",
        company="叠纸",
        sort_order=9,
        sheet_id="0RhNDB",
        spreadsheet_token="F4LLsKpfdhzxAVt8saOc26BKnLg",
    ),
    "niki_xinzuo": ProjectConfig(
        key="niki_xinzuo",
        display_name="ニキ新作",
        aliases=("ニキ新作", "X6", "x6", "niki", "niki_xinzuo"),
        template_file="104批报价单_ニキ新作_中译日模板.xlsx",
        history_dir="叠纸/ニキ新作",
        default_languages=("初翻",),
        prices={"初翻": 0.32, "初翻+审校": 0.64},
        generator="diezhi",
        file_prefix="104批报价单_ニキ新作_中译日",
        company="叠纸",
        sort_order=10,
        sheet_id="0ZUrJg",
        spreadsheet_token="PjMVsbdZvhRzOetGS7vc13Ptnrg",
    ),
    "tk": ProjectConfig(
        key="tk",
        display_name="TK项目",
        aliases=("TK", "tk", "TK项目"),
        template_file="",
        history_dir="Bilibili/TK",
        default_languages=(),
        prices={},
        generator="tk_fill",
        file_prefix="",
        company="Bilibili",
        sort_order=7,
        sheet_id="d73cd4",
        spreadsheet_token="KUTQsDzIxhg4SltVkCdceeOfnXb",
    ),
    "zulong": ProjectConfig(
        key="zulong",
        display_name="祖龙客服",
        aliases=("祖龙", "zulong", "祖龙客服"),
        template_file="",
        history_dir="祖龙",
        default_languages=(),
        prices={},
        generator="zulong_settlement",
        file_prefix="",
        company="祖龙",
        sort_order=12,
    ),
    "4399": ProjectConfig(
        key="4399",
        display_name="4399",
        aliases=("4399", "game4399"),
        template_file="",
        history_dir="4399",
        default_languages=(),
        prices={},
        generator="fill_4399",
        file_prefix="",
        company="4399",
        sort_order=11,
        sheet_id="a63cb8",
        spreadsheet_token="YcKKsgMtUhrghlt65ubczJ21n9c",
    ),
}

CONFIG_PATH = Path(__file__).parent.parent / "config" / "project_config.json"


def load_project_config():
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                configs = json.load(f)
            
            for key, config in configs.items():
                if key in PROJECTS:
                    if "display_name" in config:
                        PROJECTS[key].display_name = config["display_name"]
                    if "sort_order" in config:
                        PROJECTS[key].sort_order = config["sort_order"]
                    if "company" in config:
                        PROJECTS[key].company = config["company"]
        except Exception:
            pass


def save_project_config(project_key: str, display_name: str, sort_order: int, company: str = ""):
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    
    configs = {}
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                configs = json.load(f)
        except Exception:
            pass
    
    configs[project_key] = {
        "display_name": display_name,
        "sort_order": sort_order,
        "company": company,
    }
    
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(configs, f, ensure_ascii=False, indent=2)


def get_projects_by_company():
    load_project_config()
    
    projects_list = list(PROJECTS.values())
    projects_list.sort(key=lambda p: (p.sort_order, p.display_name))
    
    companies = {}
    for project in projects_list:
        company = project.company or "未分类"
        if company not in companies:
            companies[company] = []
        companies[company].append(project)
    
    return companies


load_project_config()


def resolve_project(name: str) -> ProjectConfig:
    normalized = name.strip().lower()
    for project in PROJECTS.values():
        if normalized == project.key.lower() or normalized in {alias.lower() for alias in project.aliases}:
            return project
    available = "、".join(project.display_name for project in PROJECTS.values())
    raise ValueError(f"未知项目: {name}。目前支持: {available}")


def find_workspace_dirs(root: Path) -> tuple[Path, Path]:
    from .paths import get_quote_history_dir
    template_dir = root / "模板" / "报价单模板"
    history_dir = get_quote_history_dir()
    try:
        template_dir.mkdir(parents=True, exist_ok=True)
        history_dir.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass
    return template_dir, history_dir
