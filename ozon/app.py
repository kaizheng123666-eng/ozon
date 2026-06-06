from __future__ import annotations

import json
import os
import re
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List
from urllib.parse import quote

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from dotenv import load_dotenv, set_key

import config
from ai_product_analyzer import PRODUCT_SCHEMA
from ai_product_analyzer import analyze_product_images, mock_product_analysis
from excel_loader import create_sample_excel, fill_ozon_template, load_products_from_excel
from image_mapper import map_product_images
from image_processor import list_product_images, process_product_images, safe_offer_id
from ozon_client import (
    OzonClient,
    products_to_import_payload,
    products_to_picture_payload,
    products_to_price_payload,
    products_to_stock_payload,
)
from result_exporter import export_results, make_result
from validators import split_valid_products, validate_products


BASE_DIR = config.BASE_DIR
INPUT_DIR = config.INPUT_DIR
IMAGE_DIR = config.IMAGE_DIR
OUTPUT_DIR = config.OUTPUT_DIR
AI_INPUT_IMAGE_DIR = getattr(config, "AI_INPUT_IMAGE_DIR", BASE_DIR / "input_images")
AI_OUTPUT_IMAGE_DIR = getattr(config, "AI_OUTPUT_IMAGE_DIR", BASE_DIR / "output_images")
AI_OUTPUT_EXCEL_DIR = getattr(config, "AI_OUTPUT_EXCEL_DIR", BASE_DIR / "output_excel")
MAOZI_PLUGIN_DIR = BASE_DIR.parent / "maozi-plugin-2.4.0" / "maozi-plugin-2.4.0"
MAOZI_PRICE_TOOL_URL = "https://ozon.maozierp.com/#/calculate"
DOWNLOADED_OZON_TEMPLATE = Path("D:/\u5c0f\u90d1\u529e\u516c\u4e13\u5c5e/\u7f51\u9875\u4e0b\u8f7d/\u5ba0\u7269\u670d\u88c5\u548c\u9774\u5b50_05.06.2026.xlsx")
PET_LIFE_JACKET_TEMPLATE = Path("D:/\u5c0f\u90d1\u529e\u516c\u4e13\u5c5e/\u7f51\u9875\u4e0b\u8f7d/\u5ba0\u7269\u6551\u751f\u8863\u6a21\u677f.xlsx")
DOG_RAINCOAT_TEMPLATE = Path("D:/\u5c0f\u90d1\u529e\u516c\u4e13\u5c5e/\u7f51\u9875\u4e0b\u8f7d/\u72d7\u72d7\u670d\u88c5 \u96e8\u8863\u6a21\u677f.xlsx")
load_settings = config.load_settings
mask_secret = config.mask_secret
ensure_directories = config.ensure_directories

for path in [AI_INPUT_IMAGE_DIR, AI_OUTPUT_IMAGE_DIR, AI_OUTPUT_EXCEL_DIR]:
    path.mkdir(parents=True, exist_ok=True)

ensure_directories()


CLOTHING_TYPE_OPTIONS = ["", "T恤", "内裤", "卫衣", "外套", "夹克", "帽子", "毛衣", "背心", "雨衣", "马甲"]
STANDARD_SIZE_OPTIONS = ["XS", "S", "M", "L", "XL", "XXL", "3XL", "4XL", "5XL", "6XL", "7XL", "8XL"]
DEFAULT_STANDARD_SIZES = ["XS", "S", "M", "L", "XL", "XXL"]
EXPECTED_PROFIT_FIELD = "期望利润CNY"

# 固定常用 Ozon 类目模板：宠物用品 / 宠物服装和靴子
# 用于页面下拉选择，避免手动输入 description_category_id 和 type_id。
PET_CATEGORY_TEMPLATE_OPTIONS = [
    {"label": "狗鞋", "description_category_id": "17028966", "type_id": "96062", "type_name": "狗鞋"},
    {"label": "宠物救生衣", "description_category_id": "17028966", "type_id": "863256801", "type_name": "宠物救生衣"},
    {"label": "宠物爪套", "description_category_id": "17028966", "type_id": "98983", "type_name": "宠物爪套"},
    {"label": "宠物饰品", "description_category_id": "17028966", "type_id": "971929672", "type_name": "宠物饰品"},
    {"label": "动物模型", "description_category_id": "17028966", "type_id": "971813657", "type_name": "动物模型"},
    {"label": "宠物服装", "description_category_id": "17028966", "type_id": "96063", "type_name": "宠物服装"},
]
PET_CATEGORY_TEMPLATE_BY_LABEL = {item["label"]: item for item in PET_CATEGORY_TEMPLATE_OPTIONS}

MODEL_PREFIX_OPTIONS = {
    "救生衣": "JSY",
    "雨衣": "YY",
    "衣服": "YF",
    "服装": "YF",
}
GENERIC_COLOR_LABELS = {"", "多色", "不同颜色", "混色", "彩色", "various", "разные цвета"}

# Ozon 标题规则：商品名称不带颜色、不带尺码；颜色写到 color / color_name_ru 字段。
OZON_TITLE_TEMPLATE_CN = "商品类型 + 款式/功能 + 适用对象 + 使用场景/核心卖点"
OZON_TITLE_TEMPLATE_RU = "Тип товара + фасон/функция + для кого + сценарий/ключевое преимущество"
OZON_TITLE_EXAMPLES_RU = [
    "Спасательный жилет для собак, регулируемый жилет для плавания и отдыха на воде",
    "Дождевик-комбинезон для собак, водонепроницаемая одежда для прогулок в дождь",
    "Обувь для собак для прогулок, защитные ботинки для лап",
]
OZON_TITLE_TEMPLATE_DETAIL = (
    "按固定顺序生成商品名称：商品类型 + 款式/功能 + 适用对象 + 使用场景/核心卖点。"
    "商品名称里禁止写颜色、尺码、SKU、货号、库存、物流词、促销词；颜色只写到 color / color_name_ru 字段，尺码只写到 sizes 字段。"
)
TITLE_COLOR_WORDS_RE = re.compile(
    r"\b(?:черн\w*|чёрн\w*|бел\w*|красн\w*|син(?:ий|яя|ее|ие|его|ему|им|ем|юю|ими|их)?|"
    r"голуб\w*|зелен\w*|зелён\w*|желт\w*|жёлт\w*|розов\w*|фиолетов\w*|"
    r"сер(?:ый|ая|ое|ые|ого|ому|ым|ом|ую|ыми|ых)?|коричнев\w*|бежев\w*|хаки|"
    r"оранжев\w*|бордов\w*|золотист\w*|серебрист\w*|прозрачн\w*|фукси\w*)\b",
    flags=re.IGNORECASE,
)
TITLE_COLOR_CN_RE = re.compile(
    r"深蓝色|浅蓝色|粉红色|咖啡色|卡其色|酒红色|玫红色|透明色|"
    r"黑色|白色|红色|蓝色|绿色|黄色|粉色|紫色|灰色|棕色|米色|橙色|金色|银色|"
    r"黑|白|红|蓝|绿|黄|粉|紫|灰|棕|米|橙"
)

COLOR_RU_MAP = {
    "黑色": "черный",
    "白色": "белый",
    "红色": "красный",
    "蓝色": "синий",
    "深蓝色": "темно-синий",
    "浅蓝色": "голубой",
    "绿色": "зеленый",
    "黄色": "желтый",
    "粉红色": "розовый",
    "紫色": "фиолетовый",
    "灰色": "серый",
    "棕色": "коричневый",
    "咖啡色": "коричневый",
    "米色": "бежевый",
    "卡其色": "хаки",
    "橙色": "оранжевый",
    "酒红色": "бордовый",
    "玫红色": "фуксия",
    "金色": "золотистый",
    "银色": "серебристый",
    "透明色": "прозрачный",
}

COLOR_CANONICAL_MAP = {
    "black": "黑色", "white": "白色", "red": "红色", "blue": "蓝色", "navy": "深蓝色",
    "green": "绿色", "yellow": "黄色", "pink": "粉红色", "purple": "紫色", "gray": "灰色",
    "grey": "灰色", "brown": "棕色", "coffee": "咖啡色", "beige": "米色", "khaki": "卡其色",
    "orange": "橙色", "gold": "金色", "silver": "银色", "transparent": "透明色",
    "черный": "黑色", "чёрный": "黑色", "белый": "白色", "красный": "红色", "синий": "蓝色",
    "голубой": "浅蓝色", "зеленый": "绿色", "зелёный": "绿色", "желтый": "黄色", "жёлтый": "黄色",
    "розовый": "粉红色", "фиолетовый": "紫色", "серый": "灰色", "коричневый": "棕色", "бежевый": "米色",
    "хаки": "卡其色", "оранжевый": "橙色", "бордовый": "酒红色", "золотистый": "金色",
    "серебристый": "银色", "прозрачный": "透明色",
    "黑": "黑色", "白": "白色", "红": "红色", "蓝": "蓝色", "绿": "绿色", "黄": "黄色", "粉": "粉红色", "粉色": "粉红色",
    "紫": "紫色", "灰": "灰色", "棕": "棕色", "咖啡": "咖啡色", "米": "米色", "卡其": "卡其色", "橙": "橙色",
}

COLOR_CODE_MAP = {
    "黑色": "HEI", "白色": "BAI", "红色": "HONG", "蓝色": "LAN", "深蓝色": "SHENLAN", "浅蓝色": "QIANLAN",
    "绿色": "LV", "黄色": "HUANG", "粉红色": "FEN", "紫色": "ZI", "灰色": "HUI", "棕色": "ZONG",
    "咖啡色": "KAFEI", "米色": "MI", "卡其色": "KAQI", "橙色": "CHENG", "酒红色": "JIUHONG",
    "玫红色": "MEIHONG", "金色": "JIN", "银色": "YIN", "透明色": "TOUMING",
}

PRODUCT_TYPE_RU_MAP = [
    ("雨衣", "Плащ-дождевик"),
    ("防水衣", "Плащ-дождевик"),
    ("救生衣", "Спасательный жилет"),
    ("外套", "Куртка"),
    ("夹克", "Куртка"),
    ("毛衣", "Свитер"),
    ("卫衣", "Худи"),
    ("背心", "Жилет"),
    ("马甲", "Жилет"),
    ("T恤", "Футболка"),
    ("连体衣", "Комбинезон"),
    ("裙子", "Платье"),
    ("衣服", "Одежда"),
    ("服装", "Одежда"),
]

PRODUCT_FEATURE_RU_MAP = [
    ("防水", "непромокаемый"),
    ("防雨", "для дождя"),
    ("反光", "со светоотражающими элементами"),
    ("可爱", "милый"),
    ("冬季", "зимний"),
    ("冬天", "зимний"),
    ("春秋", "демисезонный"),
    ("薄款", "легкий"),
    ("加厚", "утепленный"),
]

PROMO_TITLE_WORDS_RE = re.compile(
    r"\b(лучший|супер|хит|распродажа|скидка|топ|новинка|дешевый|самый|best|hot|sale|sku)\b",
    flags=re.IGNORECASE,
)
SIZE_TITLE_RE = re.compile(r"\b(?:XS|S|M|L|XL|XXL|XXXL|[2-8]XL)\b", flags=re.IGNORECASE)

TITLE_TYPE_PATTERNS = [
    (["дожд", "плащ", "雨衣", "防雨"], "плащ-дождевик"),
    (["спасател", "плаван", "救生"], "спасательный жилет"),
    (["пчел", "蜜蜂"], "костюм пчелы"),
    (["худи", "толстов", "卫衣"], "худи"),
    (["комбинез", "连体"], "комбинезон"),
    (["куртк", "外套", "夹克"], "куртка"),
    (["жилет", "马甲", "背心"], "жилет"),
    (["костюм", "造型", "cosplay"], "костюм"),
    (["衣服", "服装", "одеж"], "одежда"),
]

TITLE_STYLE_PATTERNS = [
    (["пчел", "蜜蜂"], "костюм пчелы"),
    (["полос", "条纹"], "полосатый костюм"),
    (["празд", "节日"], "праздничная одежда"),
    (["мил", "可爱"], "милый костюм"),
    (["диноз", "恐龙"], "костюм динозавра"),
    (["акул", "鲨鱼"], "костюм акулы"),
    (["медвед", "熊"], "костюм медведя"),
]

TITLE_SCENARIO_PATTERNS = [
    (["плаван", "воду", "воде", "游泳", "玩水"], "для плавания и отдыха на воде"),
    (["фото", "拍照", "фотосесс"], "для фотосессий"),
    (["празд", "节日"], "для праздников"),
    (["вечерин", "派对"], "для вечеринок"),
    (["прогул", "散步", "户外"], "для прогулок"),
    (["дожд", "雨"], "для прогулок в дождь"),
]

PET_CLOTHING_TYPE_KEYWORDS = ["одежда", "костюм", "попона", "жилет", "комбинезон", "плащ", "куртка", "худи"]

def save_uploaded_excel(uploaded_file) -> Path:
    output = INPUT_DIR / uploaded_file.name
    output.write_bytes(uploaded_file.getvalue())
    return output


def save_json(data: Dict, output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return output_path


def _clean_gpt_json_text(text: str) -> str:
    cleaned = str(text or "").strip()
    if not cleaned:
        raise ValueError("请先粘贴 GPT 返回的 JSON。")
    cleaned = cleaned.replace("\ufeff", "")
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?", "", cleaned, flags=re.IGNORECASE).strip()
        cleaned = re.sub(r"```$", "", cleaned).strip()
    cleaned = cleaned.replace("“", '"').replace("”", '"').replace("：", ":").replace("，", ",")
    match = re.search(r"\{.*\}", cleaned, flags=re.DOTALL)
    if match:
        cleaned = match.group(0)
    return cleaned


def _repair_missing_json_commas(text: str) -> str:
    lines = text.splitlines()
    repaired: List[str] = []
    for index, line in enumerate(lines):
        stripped = line.rstrip()
        content = stripped.strip()
        next_content = ""
        for next_line in lines[index + 1 :]:
            if next_line.strip():
                next_content = next_line.strip()
                break

        should_add_comma = (
            content
            and next_content.startswith('"')
            and not content.endswith((",", "{", "[", ":"))
            and (
                content.endswith('"')
                or content.endswith("}")
                or content.endswith("]")
                or content.endswith("true")
                or content.endswith("false")
                or content.endswith("null")
                or bool(re.search(r"[-\d]$", content))
            )
        )
        repaired.append(f"{stripped}," if should_add_comma else stripped)
    repaired_text = "\n".join(repaired)
    repaired_text = re.sub(r",\s*([}\]])", r"\1", repaired_text)
    return repaired_text


def extract_json_object(text: str) -> Dict[str, Any]:
    cleaned = _clean_gpt_json_text(text)
    candidates = [cleaned, _repair_missing_json_commas(cleaned)]
    last_error: Exception | None = None
    data: Any = None
    for candidate in candidates:
        try:
            data = json.loads(candidate)
            break
        except json.JSONDecodeError as exc:
            last_error = exc
    else:
        if isinstance(last_error, json.JSONDecodeError):
            raise ValueError(
                f"GPT 返回内容不是合法 JSON：第{last_error.lineno}行第{last_error.colno}列。"
                "请回到 ChatGPT，让它只返回一个完整 JSON 对象，不要解释，不要漏逗号，再粘贴回来。"
            ) from last_error
        raise ValueError("GPT 返回内容不是合法 JSON。")

    if isinstance(data, dict):
        for key in ["result", "analysis", "data", "json"]:
            value = data.get(key)
            if isinstance(value, dict):
                return value
            if isinstance(value, str):
                return extract_json_object(value)
        return data
    raise ValueError("GPT 返回内容必须是 JSON 对象。")


def get_client(client_id: str, api_key: str, base_url: str, mock_mode: bool) -> OzonClient:
    return OzonClient(client_id=client_id, api_key=api_key, base_url=base_url, mock_mode=mock_mode)


def custom_gpt_configured() -> bool:
    load_dotenv(BASE_DIR / ".env", override=True)
    return bool(os.getenv("CUSTOM_GPT_URL", "").strip() or os.getenv("GPT_PROGRAM_URL", "").strip())


def run_api_step(client: OzonClient, mode: str, products: List[Dict]) -> Dict:
    if "上传商品" in mode or "涓婁紶鍟嗗搧" in mode:
        return client.import_products(products_to_import_payload(products))
    if "上传图片" in mode or "涓婁紶鍥剧墖" in mode:
        return client.import_pictures(products_to_picture_payload(products))
    if "更新价格" in mode or "鏇存柊浠锋牸" in mode:
        return client.update_prices(products_to_price_payload(products))
    if "更新库存" in mode or "鏇存柊搴撳瓨" in mode:
        return client.update_stocks(products_to_stock_payload(products))
    if "全流程" in mode or "鍏ㄦ祦" in mode:
        product_result = client.import_products(products_to_import_payload(products))
        picture_result = client.import_pictures(products_to_picture_payload(products))
        price_result = client.update_prices(products_to_price_payload(products))
        stock_result = client.update_stocks(products_to_stock_payload(products))
        return {
            "success": all(item.get("success") for item in [product_result, picture_result, price_result, stock_result]),
            "task_id": product_result.get("task_id", ""),
            "response": {
                "product": product_result,
                "pictures": picture_result,
                "prices": price_result,
                "stocks": stock_result,
            },
        }
    return {"success": True, "response": {"message": "鍙牎楠屾ā寮忥紝娌℃湁璋冪敤 Ozon API"}}


def flatten_category_tree(data: Any) -> List[Dict[str, str]]:
    root = data.get("result", data) if isinstance(data, dict) else data
    options: List[Dict[str, str]] = []

    def walk(node: Any, path: List[str], inherited_description_category_id: Any = None) -> None:
        if isinstance(node, list):
            for item in node:
                walk(item, path, inherited_description_category_id)
            return
        if not isinstance(node, dict):
            return

        name = str(
            node.get("category_name")
            or node.get("name")
            or node.get("type_name")
            or node.get("title")
            or ""
        ).strip()
        next_path = [*path, name] if name else path
        description_category_id = node.get("description_category_id") or inherited_description_category_id
        type_id = node.get("type_id")
        if description_category_id and type_id:
            type_name = str(
                node.get("type_name")
                or node.get("name")
                or node.get("title")
                or node.get("category_name")
                or ""
            ).strip()
            category_path = " / ".join(path) or " / ".join(next_path) or f"{description_category_id}"
            reason = category_recommendation_reason(category_path, type_name)
            options.append(
                {
                    "path": category_path,
                    "label": category_path,
                    "description_category_id": str(description_category_id),
                    "type_id": str(type_id),
                    "type_name": type_name,
                    "recommendation_reason": reason,
                }
            )

        for key in ["children", "categories", "items", "types"]:
            if key in node:
                walk(node[key], next_path, description_category_id)

    walk(root, [])
    unique: Dict[tuple[str, str], Dict[str, str]] = {}
    for item in options:
        unique[(item["description_category_id"], item["type_id"])] = item
    return sorted(unique.values(), key=lambda item: (category_recommendation_rank(item), item["label"], item["type_name"]))


def category_option_label(option: Dict[str, str]) -> str:
    return (
        f"{option.get('path') or option.get('label', '')}  |  "
        f"description_category_id={option['description_category_id']}  "
        f"type_id={option['type_id']}  "
        f"type_name={option.get('type_name', '')}"
    )


def pet_category_template_label(option: Dict[str, str]) -> str:
    return (
        f"{option.get('label', option.get('type_name', ''))}  |  "
        f"description_category_id={option['description_category_id']}  "
        f"type_id={option['type_id']}"
    )


def apply_pet_category_template(option: Dict[str, str]) -> None:
    st.session_state["category_description_id"] = str(option["description_category_id"])
    st.session_state["category_type_id"] = str(option["type_id"])
    st.session_state["category_type_name"] = str(option.get("type_name", option.get("label", "")))
    st.session_state["pet_category_template_label"] = str(option.get("label", option.get("type_name", "")))


def category_recommendation_reason(path: str, type_name: str, keyword: str = "") -> str:
    text = f"{path} {type_name} {keyword}".lower()
    type_text = str(type_name or "").lower()
    matched = [word for word in PET_CLOTHING_TYPE_KEYWORDS if word in type_text]
    if matched:
        return "优先推荐：类型名称包含 " + "、".join(matched)
    if any(word in text for word in ["питом", "собак", "кошек", "живот"]):
        return "相关：宠物/动物类目"
    return "普通匹配：请结合商品实物确认"


def category_recommendation_rank(option: Dict[str, str]) -> int:
    reason = option.get("recommendation_reason", "")
    if reason.startswith("优先推荐"):
        return 0
    if reason.startswith("相关"):
        return 1
    return 2


def ozon_template_options() -> List[Path]:
    candidates: List[Path] = []
    candidates.extend([PET_LIFE_JACKET_TEMPLATE, DOG_RAINCOAT_TEMPLATE, DOWNLOADED_OZON_TEMPLATE])
    candidates.extend(sorted(INPUT_DIR.glob("*.xlsx")))
    candidates.extend(sorted((BASE_DIR.parent / "鍟嗗搧妯℃澘").glob("*.xlsx")))

    unique: List[Path] = []
    seen = set()
    for path in candidates:
        key = str(path.resolve()).lower()
        if key not in seen and path.exists():
            seen.add(key)
            unique.append(path)
    return unique


def render_category_selector(category_options: List[Dict[str, str]], category_keyword: str) -> None:
    keyword_text = category_keyword.strip().lower()
    keyword_tokens = [item.lower() for item in re.split(r"\s+", keyword_text) if item.strip()]
    if any(word in keyword_text for word in ["宠物服装", "宠物衣服", "狗衣服", "猫衣服"]):
        keyword_tokens = []

    filtered_options = [
        {
            **item,
            "recommendation_reason": category_recommendation_reason(
                item.get("path", item.get("label", "")),
                item.get("type_name", ""),
                category_keyword,
            ),
        }
        for item in category_options
        if not keyword_tokens or all(token in category_option_label(item).lower() for token in keyword_tokens)
    ]
    filtered_options = sorted(
        filtered_options,
        key=lambda item: (category_recommendation_rank(item), item.get("path", ""), item.get("type_name", "")),
    )[:300]
    if not filtered_options:
        st.info("没有匹配到类目，请换一个关键词。")
        return

    table_rows = [
        {
            "选": False,
            "类目路径 path": item.get("path", item.get("label", "")),
            "description_category_id": item.get("description_category_id", ""),
            "type_id": item.get("type_id", ""),
            "type_name / 类型名称": item.get("type_name", ""),
            "推荐匹配理由": item.get("recommendation_reason", ""),
        }
        for item in filtered_options
    ]
    edited = st.data_editor(
        pd.DataFrame(table_rows),
        use_container_width=True,
        hide_index=True,
        disabled=["类目路径 path", "description_category_id", "type_id", "type_name / 类型名称", "推荐匹配理由"],
        column_config={
            "选": st.column_config.CheckboxColumn("选"),
            "类目路径 path": st.column_config.TextColumn("类目路径 path", width="large"),
            "type_name / 类型名称": st.column_config.TextColumn("type_name / 类型名称", width="medium"),
            "推荐匹配理由": st.column_config.TextColumn("推荐匹配理由", width="medium"),
        },
        key="category_selector_table",
        height=320,
    )
    selected_rows = edited[edited["选"] == True] if "选" in edited.columns else pd.DataFrame()
    if st.button("保存选中的类目", use_container_width=True, disabled=selected_rows.empty):
        selected = selected_rows.iloc[0]
        st.session_state["category_description_id"] = str(selected["description_category_id"])
        st.session_state["category_type_id"] = str(selected["type_id"])
        st.session_state["category_type_name"] = str(selected["type_name / 类型名称"])
        if st.session_state.get("review_product"):
            st.session_state["review_product"]["description_category_id"] = st.session_state["category_description_id"]
            st.session_state["review_product"]["type_id"] = st.session_state["category_type_id"]
            st.session_state["review_product"]["type_name"] = st.session_state["category_type_name"]
        st.success(
            f"已选择 description_category_id={st.session_state['category_description_id']}，"
            f"type_id={st.session_state['category_type_id']}，"
            f"类型：{st.session_state['category_type_name']}"
        )

def load_excel_into_api_session(
    excel_path: Path,
    image_root: Path,
    dry_run: bool,
    mode: str,
    loose_draft_mode: bool = False,
) -> Dict:
    loaded = load_products_from_excel(excel_path)
    products = map_product_images(loaded["products"], image_root)
    needs_images = any(text in mode for text in ["上传商品", "上传图片", "全流程", "涓婁紶鍟嗗搧", "涓婁紶鍥剧墖", "鍏ㄦ祦"])
    validation_results = validate_products(
        products,
        require_public_image_urls=(not dry_run) and needs_images and not loose_draft_mode,
        require_image_fields=(not dry_run) and needs_images and not loose_draft_mode,
        loose_draft_mode=loose_draft_mode,
    )
    st.session_state["products"] = products
    st.session_state["validation_results"] = validation_results
    st.session_state["single_test_success"] = False
    st.session_state["loaded_api_excel_path"] = str(excel_path)
    return loaded


def refresh_api_validation(dry_run: bool, mode: str, loose_draft_mode: bool) -> None:
    products = st.session_state.get("products", [])
    if not products:
        st.session_state["validation_results"] = []
        return

    needs_images = any(text in mode for text in ["上传商品", "上传图片", "全流程", "涓婁紶鍟嗗搧", "涓婁紶鍥剧墖", "鍏ㄦ祦"])
    st.session_state["validation_results"] = validate_products(
        products,
        require_public_image_urls=(not dry_run) and needs_images and not loose_draft_mode,
        require_image_fields=(not dry_run) and needs_images and not loose_draft_mode,
        loose_draft_mode=loose_draft_mode,
    )


TARGET_PET_CN_OPTIONS = ["对于狗", "对于鸟类", "猫咪用品"]


def target_pet_cn_from_context(*values: Any) -> str:
    text = " ".join(str(value or "") for value in values).lower()
    if any(token in text for token in ["鸟", "鹦鹉", "bird", "птиц", "попуг"]):
        return "对于鸟类"
    if any(token in text for token in ["猫", "cat", "коше", "кошк", "кот"]):
        return "猫咪用品"
    return "对于狗"


def build_review_product(analysis: Dict, processed: Dict, defaults: Dict) -> Dict:
    offer_id = safe_offer_id(defaults["offer_id"])
    name_ru = str(defaults.get("name_ru") or analysis.get("name_ru", "")).strip()
    product_name_cn = str(defaults.get("product_name_cn") or "").strip()
    manual_main_color = normalize_color_cn(defaults.get("manual_main_color", ""))
    color = manual_main_color or normalize_color_cn(analysis.get("color", ""))
    analysis_color_variants_text = format_color_variants(analysis)
    primary_image, additional_images = build_image_links(processed, defaults)
    draft_product = {
        "product_name_cn": product_name_cn,
        "name_ru": name_ru or translate_product_name_to_ru(product_name_cn),
        "category_name": defaults.get("category_name", ""),
        "type": "宠物服装",
        "target_pet": analysis.get("target_audience", ""),
        "style": analysis.get("style", ""),
        "season": analysis.get("season", ""),
        "color": color,
        "color_name_ru": analysis.get("color_name_ru", ""),
        "clothing_type": str(defaults.get("clothing_type", "")).strip(),
        "brand": "无品牌",
    }
    
    name_ru = generate_ozon_title_ru(draft_product, analysis)

    description_ru = (
        str(analysis.get("description_ru", "") or "").strip()
        or make_description_ru(name_ru)
    )

    return {        "offer_id": offer_id,
        "product_name_cn": product_name_cn,
        "name_ru": name_ru,
        "description_ru": description_ru,
        "description_category_id": str(defaults.get("description_category_id", "")).strip(),
        "type_id": str(defaults.get("type_id", "")).strip(),
        "price": defaults.get("price", ""),
        "old_price": defaults.get("old_price", ""),
        "stock": defaults.get("stock", 100),
        "weight": defaults.get("weight", ""),
        "product_weight": defaults.get("product_weight", ""),
        "depth": defaults.get("depth", ""),
        "width": defaults.get("width", ""),
        "height": defaults.get("height", ""),
        "primary_image": primary_image,
        "additional_images": additional_images,
        "brand": "无品牌",
        "model_prefix": "",
        "model": str(defaults.get("model", "")).strip(),
        "animal_gender": "男女两用的",
        "color": color,
        "color_name_ru": color or guess_color_name_ru(analysis.get("color", ""), analysis.get("color_name_ru", "")),
        "color_variants_text": analysis_color_variants_text or manual_main_color,
        "size_variants_text": format_size_variants(analysis.get("sizes", [])),
        "type": "宠物服装",
        "target_pet": target_pet_cn_from_context(
            defaults.get("product_name_cn", ""),
            analysis.get("target_audience", ""),
            analysis.get("product_type", ""),
            analysis.get("product_type_ru", ""),
            analysis.get("style", ""),
        ),
        "clothing_type": str(defaults.get("clothing_type", "")).strip(),
        "standard_sizes_text": str(defaults.get("standard_sizes_text", "")).strip(),
        "unit_product_quantity": 1,
        "items_in_product": 1,
        "product_type_ru": analysis.get("product_type_ru", ""),
        "tags_ru": ensure_search_tags_ru(
            {
                "product_name_cn": product_name_cn,
                "name_ru": name_ru,
                "product_type": analysis.get("product_type", ""),
                "product_type_ru": analysis.get("product_type_ru", ""),
                "clothing_type": str(defaults.get("clothing_type", "")).strip(),
                "style": analysis.get("style", ""),
                "season": analysis.get("season", ""),
                "search_keywords_ru": analysis.get("search_keywords_ru", []),
            },
            analysis.get("tags_ru", []),
        ),
        "material": analysis.get("material", ""),
        "brand_country": "中国",
        "origin_country": "中国",
        "ai_analysis": analysis,
    }


def split_variant_items(text: str) -> List[str]:
    raw = str(text or "").strip()
    if not raw:
        return []
    items: List[str] = []
    for line in raw.splitlines():
        line = line.strip()
        if not line:
            continue
        if "|" in line or "\t" in line:
            items.append(line)
        else:
            items.extend(part.strip() for part in re.split(r"[锛?锛?]+", line) if part.strip())
    return items


def split_variant_fields(item: str) -> List[str]:
    return [part.strip() for part in re.split(r"\s*\|\s*|\t+", str(item or "").strip())]


def normalize_color_cn(color: str) -> str:
    text = str(color or "").strip()
    if not text:
        return ""

    lowered = text.lower().replace("褢", "械")
    for key in sorted(COLOR_CANONICAL_MAP, key=len, reverse=True):
        lookup = key.lower().replace("褢", "械")
        if lookup and lookup in lowered:
            return COLOR_CANONICAL_MAP[key]
    return text


def color_code(value: str, fallback: str) -> str:
    normalized = normalize_color_cn(value)
    if normalized in COLOR_CODE_MAP:
        return COLOR_CODE_MAP[normalized]
    return offer_suffix(normalized, fallback)


def normalize_size_label(value: str) -> str:
    text = str(value or "").strip().upper().replace(" ", "")
    aliases = {"2XL": "XXL"}
    return aliases.get(text, text)


def looks_like_size_label(value: str) -> bool:
    text = normalize_size_label(value)
    return bool(re.fullmatch(r"(?:X{0,3}S|M|L|X{1,3}L|[3-9]XL|\d+XL)", text))


def size_sort_key(value: str) -> tuple[int, str]:
    label = normalize_size_label(value)
    order = {size: idx for idx, size in enumerate(STANDARD_SIZE_OPTIONS)}
    return order.get(label, 100), label


def parse_standard_sizes(text: str) -> List[str]:
    labels: List[str] = []
    for item in re.split(r"[\s,锛?锛泑/]+", str(text or "")):
        label = normalize_size_label(item)
        if label and looks_like_size_label(label) and label not in labels:
            labels.append(label)
    return labels


def public_image_url(file_name: str, base_url: str = "") -> str:
    text = str(file_name or "").strip()
    if not text:
        return ""
    if re.match(r"^https?://", text, flags=re.IGNORECASE):
        return text
    base = str(base_url or "").strip()
    if not base:
        return ""
    return f"{base.rstrip('/')}/{quote(Path(text).name)}"


def build_image_links(processed: Dict, defaults: Dict) -> tuple[str, List[str]]:
    main_url = str(defaults.get("main_image_url", "")).strip()
    extra_urls = [
        item.strip()
        for item in str(defaults.get("additional_image_urls", "")).replace(",", "\n").splitlines()
        if item.strip()
    ]
    if main_url or extra_urls:
        return main_url, extra_urls

    image_base_url = str(defaults.get("image_base_url", "")).strip()
    primary = public_image_url(processed.get("primary_image", ""), image_base_url)
    extras = [public_image_url(item, image_base_url) for item in processed.get("additional_images", [])]
    return primary, [item for item in extras if item]


def translate_product_name_to_ru(name_cn: str, fallback: str = "") -> str:
    text = str(name_cn or "").strip()
    if not text:
        return str(fallback or "").strip()

    lowered = text.lower()
    product_type = ""
    for key, value in PRODUCT_TYPE_RU_MAP:
        if key.lower() in lowered or key in text:
            product_type = value
            break
    if not product_type:
        product_type = "Товар для питомцев"

    if any(key in text for key in ["狗", "犬", "幼犬", "小狗"]):
        target = "для собак"
    elif any(key in text for key in ["猫", "小猫"]):
        target = "для кошек"
    elif "宠物" in text:
        target = "для питомцев"
    else:
        target = "для питомцев"

    features = []
    for key, value in PRODUCT_FEATURE_RU_MAP:
        if key in text and value not in features:
            features.append(value)
    parts = [product_type, target, *features]
    return " ".join(part for part in parts if part).strip()


def make_description_ru(name_ru: str) -> str:
    name = str(name_ru or "").strip() or "Товар для питомцев"
    return f"{name}. Перед покупкой проверьте размер по таблице."


def _context_text(*values: Any) -> str:
    return " ".join(str(value or "") for value in values).lower()


def _first_pattern_value(text: str, patterns: List[tuple[List[str], str]], default: str) -> str:
    for keys, value in patterns:
        if any(key.lower() in text for key in keys):
            return value
    return default


def _title_target_pet(context: str) -> str:
    has_cat = any(key in context for key in ["коше", "кошк", "кот", "猫"])
    has_puppy = any(key in context for key in ["щен", "幼犬", "小狗"])
    has_small_dog = any(key in context for key in ["маленьк", "мелк", "小型犬", "小狗"])
    has_dog = any(key in context for key in ["собак", "пес", "狗", "犬"])
    if has_cat and has_puppy:
        return "для кошек и щенков"
    if has_cat and has_dog:
        return "для кошек и собак"
    if has_cat:
        return "для кошек"
    if has_puppy:
        return "для щенков"
    if has_small_dog:
        return "для маленьких собак"
    if has_dog:
        return "для собак"
    return "для питомцев"


def remove_color_words_from_title(title: str) -> str:
    """商品标题不带颜色；颜色只写到 color / color_name_ru 字段。"""
    text = str(title or "").strip()
    if not text:
        return ""
    text = re.sub(r"\b(?:цвет|цвета)\s*[:：\-]?\s*", " ", text, flags=re.IGNORECASE)
    text = TITLE_COLOR_WORDS_RE.sub(" ", text)
    text = TITLE_COLOR_CN_RE.sub(" ", text)
    return text


def clean_ozon_title_ru(title: str) -> str:
    text = str(title or "").strip()
    text = re.sub(r"[#🔥⭐✨✅🚚🎉]+", " ", text)
    text = PROMO_TITLE_WORDS_RE.sub("", text)
    text = remove_color_words_from_title(text)
    text = re.sub(r"\bSKU[-A-ZА-Я0-9]*\b", "", text, flags=re.IGNORECASE)
    text = SIZE_TITLE_RE.sub("", text)
    text = re.sub(r"\b\d{2,}[-A-ZА-Я0-9]*\b", "", text)
    text = re.sub(r"\s*[,;/|]+\s*", ", ", text)
    text = re.sub(r"\s+", " ", text).strip(" ,.-")
    if len(text) > 255:
        text = text[:255].rsplit(" ", 1)[0].strip(" ,.-")
    return text[0].upper() + text[1:] if text else ""


def generate_ozon_title_ru(product: Dict, analysis: Dict | None = None) -> str:
    analysis = analysis or product.get("ai_analysis", {}) or {}
    context = _context_text(
        product.get("product_name_cn"),
        product.get("name_ru"),
        product.get("category_name"),
        product.get("type"),
        product.get("clothing_type"),
        product.get("target_pet"),
        product.get("style"),
        product.get("season"),
        analysis.get("product_type"),
        analysis.get("product_type_ru"),
        analysis.get("target_audience"),
        analysis.get("style"),
        analysis.get("season"),
        " ".join(analysis.get("main_features", []) or []),
    )
    product_type = _first_pattern_value(context, TITLE_TYPE_PATTERNS, "одежда")
    style = _first_pattern_value(context, TITLE_STYLE_PATTERNS, "")
    target = _title_target_pet(context)
    scenario = _first_pattern_value(context, TITLE_SCENARIO_PATTERNS, "")
    if any(key in context for key in ["пчел", "蜜蜂", "празд", "节日", "вечерин", "派对", "фото", "拍照"]):
        scenario = "для праздников и фотосессий"
    core = style or product_type
    if style and product_type not in style and product_type != "одежда":
        core = f"{product_type} {style}"
    brand = str(product.get("brand", "")).strip()
    if brand in {"", "无品牌", "No brand", "Нет бренда"}:
        brand = ""
    title = " ".join(part for part in [brand, core, target] if part)
    if scenario:
        title = f"{title}, {scenario}"
    if len(title) < 60 and product_type == "плащ-дождевик":
        title = f"{title}, одежда для прогулок в дождь"
    if len(title) < 60 and product_type == "спасательный жилет" and "для плавания" not in title:
        title = f"{title}, для плавания и отдыха на воде"
    return clean_ozon_title_ru(title)


def normalize_search_tag(value: Any) -> str:
    text = str(value or "").strip().lower()
    if not text:
        return ""
    text = text.lstrip("#").replace("_", " ")
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return ""
    return "#" + text.replace(" ", "_")


def russian_search_tag_candidates(product: Dict) -> List[str]:
    text = " ".join(
        str(product.get(field, ""))
        for field in ["product_name_cn", "name_ru", "product_type", "product_type_ru", "clothing_type", "style", "season"]
    ).lower()
    is_raincoat = any(token in text for token in ["雨衣", "дожд", "плащ"])
    is_life_jacket = any(token in text for token in ["救生", "спасател", "жилет"])
    is_cat = any(token in text for token in ["猫", "коше", "кошк"])

    target = "кошек" if is_cat else "собак"
    animal = "кошки" if is_cat else "собаки"

    if is_life_jacket:
        primary = [
            f"спасательный жилет для {target}",
            f"жилет для плавания {target}",
            f"плавательный жилет для {target}",
            f"жилет для {target} на воду",
            f"одежда для {target} для плавания",
            f"товары для {target} на лето",
            f"аксессуары для {target}",
        ]
    elif is_raincoat:
        primary = [
            f"дождевик для {target}",
            f"плащ для {target}",
            f"одежда для {target} от дождя",
            f"непромокаемая одежда для {target}",
            f"куртка дождевик для {target}",
            f"защита от дождя для {target}",
            f"одежда для {target} на прогулку",
        ]
    else:
        primary = [
            f"одежда для {target}",
            f"костюм для {target}",
            f"куртка для {target}",
            f"жилет для {target}",
            f"одежда для {target} на прогулку",
            f"товары для {target}",
            f"аксессуары для {target}",
        ]

    common = [
        f"одежда для маленьких {target}",
        f"одежда для средних {target}",
        f"одежда для крупных {target}",
        f"удобная одежда для {target}",
        f"легкая одежда для {target}",
        f"одежда для {target} весна",
        f"одежда для {target} осень",
        f"одежда для {target} лето",
        f"одежда для {target} на улицу",
        f"одежда для {animal}",
        "зоотовары",
        "товары для питомцев",
        "товары для животных",
        f"для {target}",
        f"{animal} одежда",
        f"{animal} прогулка",
        f"{animal} лето",
        f"{animal} осень",
        f"{animal} весна",
        f"{animal} дождь",
        f"{animal} улица",
        f"{animal} отдых",
        f"{animal} фото",
        f"костюм питомца",
        f"одежда питомца",
        f"жилет питомца",
        f"плащ питомца",
        f"куртка питомца",
        f"комбинезон питомца",
        f"худи питомца",
        f"маленькие {target}",
        f"средние {target}",
        f"крупные {target}",
        "одежда для питомцев",
        "аксессуары питомцам",
        "прогулка с собакой",
        "питомцы одежда",
    ]
    return primary + common


def ensure_search_tags_ru(product: Dict, tags: List[Any] | None = None, min_count: int = 30, max_count: int = 30) -> List[str]:
    merged = list(tags or []) + list(product.get("search_keywords_ru") or []) + russian_search_tag_candidates(product)
    result: List[str] = []
    seen = set()
    for item in merged:
        tag = normalize_search_tag(item)
        if not tag or tag in seen or len(tag) > 30:
            continue
        seen.add(tag)
        result.append(tag)
        if len(result) >= max_count:
            break
    if len(result) < min_count:
        for item in [
            "одежда для собак",
            "дождевик для собак",
            "плащ для собак",
            "жилет для собак",
            "костюм для собак",
            "куртка для собак",
            "комбинезон собаке",
            "худи для собак",
            "товары для собак",
            "товары питомцам",
            "для маленьких собак",
            "для средних собак",
            "для крупных собак",
            "одежда кошкам",
            "костюм для кошек",
            "жилет для кошек",
            "плащ для кошек",
            "для щенков",
            "для котов",
            "для питомцев",
            "прогулка",
            "фотосессия",
            "праздник",
            "вечеринка",
            "лето",
            "осень",
            "весна",
            "дождь",
            "плавание",
            "зоотовары",
        ]:
            tag = normalize_search_tag(item)
            if tag and tag not in seen and len(tag) <= 30:
                seen.add(tag)
                result.append(tag)
            if len(result) >= min_count:
                break
    return result[:max_count]


def build_gpt_plus_prompt(
    image_names: List[str],
    product_name_cn: str = "",
    colors_text: str = "",
    sizes_text: str = "",
    title_template_cn: str = OZON_TITLE_TEMPLATE_CN,
    title_template_ru: str = OZON_TITLE_TEMPLATE_RU,
) -> str:
    title_template_cn = str(title_template_cn or OZON_TITLE_TEMPLATE_CN).strip()
    title_template_ru = str(title_template_ru or OZON_TITLE_TEMPLATE_RU).strip()
    helper_data = {
        "product_name_cn": product_name_cn,
        "title_template_cn": title_template_cn,
        "title_template_ru": title_template_ru,
        "known_colors": colors_text,
        "known_sizes": sizes_text,
        "image_names": image_names,
    }
    return (
        "你是 Ozon 俄罗斯电商宠物服装商品资料整理助手。请根据我上传的商品图片识别商品资料，"
        "并且只返回一个合法 JSON 对象，不要解释，不要 Markdown。\n\n"
        "重点要求：\n"
        "1. 商品名称、简介、标签必须输出俄语。\n"
        f"2. name_ru 必须是 1 个最佳 Ozon 风格俄语标题，必须严格按这个模板顺序生成：{title_template_cn}。\n"
        f"   俄语模板参考：{title_template_ru}。\n"
        "3. 标题不要写颜色、尺码、SKU、成本、库存、物流词、爆款、热卖、跨境专用、лучший、супер、хит、распродажа、表情符号或无关关键词。颜色必须只写到 color 和 color_name_ru 字段。\n"
        "4. 标题控制在 60-120 个俄文字符左右，最长不超过 255 个字符。\n"
        "5. 尺码放到 sizes 字段，不要放进标题；如果有 XS/S/M/L/XL/XXL 等尺码，不要遗漏。\n"
        "6. 如果同一商品有多个颜色，把 sizes 当作整组商品的统一尺码池，不要让某个颜色漏尺码。\n"
        "7. color 只填商品主色中文，例如绿色、黑色、红色；不要写恐龙图案、卡通绿色等款式描述。\n"
        "8. 看不清或图片没有的信息留空，不要编造。\n"
        "9. tags_ru 需要生成 30 个，并以 # 开头；单个标签长度不要超过 30 个字符；使用自然俄语搜索短语，覆盖品类、宠物对象、用途、季节、场景和同义词。\n\n"
        "人工已知信息如下：\n"
        + json.dumps(helper_data, ensure_ascii=False, indent=2)
        + "\n\n必须严格返回符合这个 JSON schema 的对象：\n"
        + json.dumps(PRODUCT_SCHEMA, ensure_ascii=False, indent=2)
    )


def render_copy_prompt(prompt: str) -> None:
    prompt_js = json.dumps(prompt, ensure_ascii=True).replace("</", "<\\/")
    components.html(
        f'''
        <div style="font-family: sans-serif; display: flex; align-items: center; gap: 10px;">
          <button id="copy-gpt-prompt" style="padding: 8px 12px; border: 1px solid #ccc; border-radius: 6px; background: #fff; cursor: pointer;">
            复制 GPT 提示词
          </button>
          <span id="copy-status" style="color: #0f766e;"></span>
        </div>
        <textarea id="copy-source" style="position: fixed; left: -9999px; top: -9999px;"> </textarea>
        <script>
          const promptText = {prompt_js};
          const button = document.getElementById("copy-gpt-prompt");
          const status = document.getElementById("copy-status");
          const area = document.getElementById("copy-source");
          button.addEventListener("click", async () => {{
            try {{
              if (navigator.clipboard && window.isSecureContext) {{
                await navigator.clipboard.writeText(promptText);
              }} else {{
                area.value = promptText;
                area.focus();
                area.select();
                document.execCommand("copy");
              }}
              status.textContent = "已复制";
            }} catch (err) {{
              area.value = promptText;
              area.focus();
              area.select();
              const ok = document.execCommand("copy");
              status.textContent = ok ? "已复制" : "复制失败，请手动复制下方提示词";
            }}
          }});
        </script>
        ''',
        height=56,
    )


def guess_color_name_ru(color: str, current: str = "") -> str:
    current_text = str(current or "").strip()
    if current_text:
        return normalize_color_cn(current_text)

    return normalize_color_cn(color)


def format_color_variants(analysis: Dict) -> str:
    lines: List[str] = []
    for item in analysis.get("color_variants", []) or []:
        if isinstance(item, dict):
            color = normalize_color_cn(str(item.get("color", "")).strip())
            color_name = color or guess_color_name_ru(color, str(item.get("color_name_ru", "")).strip())
            if color or color_name:
                lines.append("|".join([color, color_name]).rstrip("|"))
    if not lines:
        color = normalize_color_cn(str(analysis.get("color", "")).strip())
        color_name = color or guess_color_name_ru(color, str(analysis.get("color_name_ru", "")).strip())
        if color or color_name:
            lines.append("|".join([color, color_name]).rstrip("|"))
    return "\n".join(lines)


def format_size_variants(sizes: List[Dict]) -> str:
    lines: List[str] = []
    for item in sizes or []:
        if isinstance(item, str):
            if item.strip():
                lines.append(item.strip())
            continue
        if not isinstance(item, dict):
            continue
        fields = [
            str(item.get("size", "")).strip(),
            str(item.get("chest_cm", "")).strip(),
            str(item.get("back_length_cm", "")).strip(),
            str(item.get("neck_cm", "")).strip(),
            str(item.get("price", "")).strip(),
            str(item.get("weight", "")).strip(),
            str(item.get("depth", "")).strip(),
            str(item.get("width", "")).strip(),
            str(item.get("height", "")).strip(),
        ]
        if not fields[0]:
            continue
        if any(fields[1:]):
            lines.append("|".join(fields).rstrip("|"))
        else:
            lines.append(fields[0])
    return "\n".join(lines)


def first_number(value: str) -> str:
    match = re.search(r"\d+(?:[.,]\d+)?", str(value or ""))
    return match.group(0).replace(",", ".") if match else ""


def parse_measure_range(value: str) -> tuple[str, str]:
    text = str(value or "").strip()
    if not text:
        return "", ""
    text = re.sub(r"(厘米|公分|cm|CM|胸围|胸|颈围|脖围|脖子|脖)", "", text).strip()
    numbers = re.findall(r"\d+(?:[.,]\d+)?", text)
    if not numbers:
        return "", ""
    if len(numbers) == 1:
        value_text = numbers[0].replace(",", ".")
        return "", value_text
    first = numbers[0].replace(",", ".")
    second = numbers[1].replace(",", ".")
    try:
        low, high = sorted([float(first), float(second)])
        return f"{low:g}", f"{high:g}"
    except ValueError:
        return first, second


def parse_package_dimensions(value: str, default_depth: int = 300, default_width: int = 220, default_height: int = 30) -> tuple[int, int, int]:
    numbers = re.findall(r"\d+(?:[.,]\d+)?", str(value or ""))
    if len(numbers) < 3:
        return default_depth, default_width, default_height
    try:
        depth, width, height = [int(float(item.replace(",", "."))) for item in numbers[:3]]
        return depth, width, height
    except ValueError:
        return default_depth, default_width, default_height


def looks_like_old_size_format(fields: List[str]) -> bool:
    if len(fields) != 6:
        return False
    numeric_tail = [first_number(item) for item in fields[3:5]]
    try:
        return len(numeric_tail) == 2 and all(float(item) >= 100 for item in numeric_tail)
    except ValueError:
        return False


def parse_color_variants(text: str, fallback_color: str, fallback_color_name: str) -> List[Dict]:
    variants: List[Dict] = []
    for item in split_variant_items(text):
        fields = split_variant_fields(item)
        color = normalize_color_cn(fields[0] if fields else "")
        color_name = guess_color_name_ru(color, fields[1] if len(fields) > 1 else "")
        if color.lower() in {"color", "颜色", "商品颜色"} or color in GENERIC_COLOR_LABELS:
            continue
        if color or color_name:
            variants.append({"color": color, "color_name_ru": color_name})
    if variants:
        return variants
    fallback = normalize_color_cn(fallback_color)
    if fallback in GENERIC_COLOR_LABELS:
        return [{"color": "", "color_name_ru": ""}]
    return [{"color": fallback, "color_name_ru": guess_color_name_ru(fallback, fallback_color_name)}]


def parse_size_variants(text: str) -> List[Dict]:
    variants: List[Dict] = []
    for item in split_variant_items(text):
        fields = split_variant_fields(item)
        if not fields:
            continue
        if len(fields) > 1 and not looks_like_size_label(fields[0]) and looks_like_size_label(fields[1]):
            fields = fields[1:]
        size = normalize_size_label(fields[0])
        if not size or size.lower() in {"size", "尺码", "尺寸"}:
            continue

        chest_min, chest_max = parse_measure_range(fields[1] if len(fields) > 1 else "")
        back_length = first_number(fields[2] if len(fields) > 2 else "")
        neck_min, neck_max = parse_measure_range(fields[3] if len(fields) > 3 else "")
        price_index = 4

        variants.append(
            {
                "pet_size": size,
                "chest_min_cm": chest_min,
                "chest_max_cm": chest_max,
                "back_length_cm": back_length,
                "neck_min_cm": neck_min,
                "neck_max_cm": neck_max,
                "price": fields[price_index] if len(fields) > price_index else "",
                "weight": fields[price_index + 1] if len(fields) > price_index + 1 else "",
                "depth": fields[price_index + 2] if len(fields) > price_index + 2 else "",
                "width": fields[price_index + 3] if len(fields) > price_index + 3 else "",
                "height": fields[price_index + 4] if len(fields) > price_index + 4 else "",
            }
        )
    deduped: Dict[str, Dict] = {}
    for variant in variants:
        size_key = normalize_size_label(variant.get("pet_size", ""))
        if not size_key:
            continue
        if size_key not in deduped:
            deduped[size_key] = variant
            continue
        current = deduped[size_key]
        for field, value in variant.items():
            if field == "pet_size":
                continue
            if str(value or "").strip() and not str(current.get(field, "")).strip():
                current[field] = value

    return sorted(deduped.values(), key=lambda item: size_sort_key(item.get("pet_size", ""))) or [
        {
            "pet_size": "",
            "chest_min_cm": "",
            "chest_max_cm": "",
            "back_length_cm": "",
            "neck_min_cm": "",
            "neck_max_cm": "",
            "price": "",
            "weight": "",
            "depth": "",
            "width": "",
            "height": "",
        }
    ]


def empty_size_variant(size: str) -> Dict:
    return {
        "pet_size": normalize_size_label(size),
        "chest_min_cm": "",
        "chest_max_cm": "",
        "back_length_cm": "",
        "neck_min_cm": "",
        "neck_max_cm": "",
        "price": "",
        "weight": "",
        "depth": "",
        "width": "",
        "height": "",
    }


def ensure_size_pool(sizes: List[Dict], standard_sizes_text: str = "") -> List[Dict]:
    real_sizes = [item for item in sizes if str(item.get("pet_size", "")).strip()]
    if not real_sizes:
        real_sizes = []

    by_size: Dict[str, Dict] = {normalize_size_label(item.get("pet_size", "")): dict(item) for item in real_sizes}
    for label in parse_standard_sizes(standard_sizes_text):
        by_size.setdefault(label, empty_size_variant(label))

    if not by_size:
        return sizes
    return sorted(by_size.values(), key=lambda item: size_sort_key(item.get("pet_size", "")))


def offer_suffix(value: str, fallback: str) -> str:
    ascii_text = re.sub(r"[^A-Za-z0-9]+", "-", str(value or "").strip()).strip("-").upper()
    return ascii_text or fallback


def infer_model_prefix(product_name: str, clothing_type: str = "") -> str:
    text = f"{product_name or ''} {clothing_type or ''}"
    for keyword, prefix in MODEL_PREFIX_OPTIONS.items():
        if keyword in text:
            return prefix
    return "YF"


def build_model_card_name(product: Dict) -> str:
    manual_model = str(product.get("model", "")).strip()
    if manual_model:
        return manual_model
    product_name = str(product.get("product_name_cn") or product.get("name_ru", "")).strip()
    prefix = str(product.get("model_prefix") or "").strip().upper()
    if not prefix:
        prefix = infer_model_prefix(product_name, str(product.get("clothing_type", "")))
    if product_name.upper().startswith(f"{prefix}-"):
        return product_name
    if product_name:
        return f"{prefix}-{product_name}"
    return f"{prefix}-"


def build_variant_products(product: Dict) -> List[Dict]:
    colors = parse_color_variants(
        product.get("color_variants_text", ""),
        str(product.get("color", "")),
        str(product.get("color_name_ru", "")),
    )
    sizes = ensure_size_pool(
        parse_size_variants(product.get("size_variants_text", "")),
        str(product.get("standard_sizes_text", "")),
    )
    color_is_variant = len(colors) > 1
    size_is_variant = len(sizes) > 1 or any(str(item.get("pet_size", "")).strip() for item in sizes)

    products: List[Dict] = []
    base_offer_id = safe_offer_id(product.get("offer_id", ""))
    base_model = build_model_card_name(product)
    color_code_counts: Dict[str, int] = {}
    for color_index, color in enumerate(colors, start=1):
        current_color = normalize_color_cn(color.get("color", ""))
        current_color_code = color_code(current_color, f"C{color_index:02d}") if current_color else ""
        color_code_counts[current_color_code] = color_code_counts.get(current_color_code, 0) + 1
        color_suffix = current_color_code
        if current_color_code and color_code_counts[current_color_code] > 1:
            color_suffix = f"{current_color_code}{color_code_counts[current_color_code]}"
        for size_index, size in enumerate(sizes, start=1):
            row = deepcopy(product)
            row["color"] = current_color
            row["color_name_ru"] = guess_color_name_ru(row["color"], color.get("color_name_ru", ""))
            row["pet_size"] = size.get("pet_size", "")
            row["chest_min_cm"] = size.get("chest_min_cm", "")
            row["chest_max_cm"] = size.get("chest_max_cm", "")
            row["back_length_cm"] = size.get("back_length_cm", "")
            row["neck_min_cm"] = size.get("neck_min_cm", "")
            row["neck_max_cm"] = size.get("neck_max_cm", "")
            row["model"] = base_model
            row["unit_product_quantity"] = product.get("unit_product_quantity", 1) or 1
            row["items_in_product"] = product.get("items_in_product", 1) or 1
            for field in ["price", "weight", "depth", "width", "height"]:
                if str(size.get(field, "")).strip():
                    row[field] = size[field]

            suffixes: List[str] = []
            if color_suffix:
                suffixes.append(color_suffix)
            if size_is_variant:
                suffixes.append(offer_suffix(size.get("pet_size", ""), f"S{size_index:02d}"))
            row["offer_id"] = safe_offer_id("-".join([base_offer_id, *suffixes]) if suffixes else base_offer_id)
            products.append(row)
    return products


def number_value(value: Any, default: float = 0.0) -> float:
    text = str(value if value is not None else "").strip().replace(",", ".")
    if not text:
        return default
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return default
    try:
        return float(match.group(0))
    except ValueError:
        return default


def short_text(value: Any, limit: int = 80) -> str:
    text = str(value or "").strip()
    if len(text) <= limit:
        return text
    return text[: limit - 1] + "…"


def round_up_to_step(value: float, step: float) -> float:
    if step <= 0:
        return round(value, 2)
    rounded = ((int((value + step - 0.000001) / step)) * step)
    return round(rounded, 2)


def row_commission_percent(row: Dict, default_commission_percent: float) -> float:
    return number_value(row.get("平台抽佣%"), default_commission_percent)


def suggested_price_cny(row: Dict, commission_percent: float, profit_margin_percent: float, round_step: float) -> float:
    base_cost = (
        number_value(row.get("成本价CNY"))
        + number_value(row.get("国内运费CNY"))
        + number_value(row.get("跨境物流CNY"))
        + number_value(row.get("包装费CNY"))
        + number_value(row.get("其他成本CNY"))
    )
    expected_profit = number_value(row.get(EXPECTED_PROFIT_FIELD), 0.0)
    if expected_profit > 0:
        denominator = 1 - commission_percent / 100
        if base_cost <= 0 or denominator <= 0:
            return 0.0
        return round_up_to_step((base_cost + expected_profit) / denominator, round_step)

    denominator = 1 - commission_percent / 100 - profit_margin_percent / 100
    if base_cost <= 0 or denominator <= 0:
        return 0.0
    return round_up_to_step(base_cost / denominator, round_step)


def update_pricing_calculations(
    rows: List[Dict],
    default_commission_percent: float,
    profit_margin_percent: float,
    round_step: float,
) -> List[Dict]:
    calculated: List[Dict] = []
    for row in rows:
        next_row = dict(row)
        commission_percent = row_commission_percent(next_row, default_commission_percent)
        manual_price = number_value(next_row.get("手动售价CNY"), 0.0)
        suggested_price = manual_price or suggested_price_cny(
            next_row,
            commission_percent,
            profit_margin_percent,
            round_step,
        )
        next_row["平台抽佣CNY"] = round(suggested_price * commission_percent / 100, 2) if suggested_price > 0 else ""
        next_row["建议售价CNY"] = suggested_price if suggested_price > 0 else ""
        calculated.append(next_row)
    return calculated


def build_pricing_rows(products: List[Dict], existing_rows: List[Dict], default_commission_percent: float) -> List[Dict]:
    existing_by_offer = {str(row.get("offer_id", "")): row for row in existing_rows if str(row.get("offer_id", "")).strip()}
    rows: List[Dict] = []
    for product in products:
        offer_id = str(product.get("offer_id", ""))
        existing = existing_by_offer.get(offer_id, {})
        existing_commission = existing.get("平台抽佣%", "")
        rows.append(
            {
                "offer_id": offer_id,
                "颜色": product.get("color", ""),
                "尺码": product.get("pet_size", ""),
                "成本价CNY": existing.get("成本价CNY", ""),
                "国内运费CNY": existing.get("国内运费CNY", ""),
                "跨境物流CNY": existing.get("跨境物流CNY", ""),
                "包装费CNY": existing.get("包装费CNY", ""),
                "其他成本CNY": existing.get("其他成本CNY", ""),
                "平台抽佣%": existing_commission if str(existing_commission).strip() else default_commission_percent,
                "平台抽佣CNY": existing.get("平台抽佣CNY", ""),
                "手动售价CNY": existing.get("手动售价CNY", ""),
                "建议售价CNY": existing.get("建议售价CNY", product.get("price", "")),
            }
        )
    for row in rows:
        row.setdefault(EXPECTED_PROFIT_FIELD, "")
    return rows


def apply_pricing_to_products(products: List[Dict], pricing_rows: List[Dict]) -> List[Dict]:
    price_by_offer = {
        str(row.get("offer_id", "")): row.get("建议售价CNY", "")
        for row in pricing_rows
        if str(row.get("offer_id", "")).strip() and str(row.get("建议售价CNY", "")).strip()
    }
    updated: List[Dict] = []
    for product in products:
        row = dict(product)
        price = price_by_offer.get(str(product.get("offer_id", "")))
        if price is not None and str(price).strip():
            row["price"] = price
        updated.append(row)
    return updated


def render_pricing_tool(products: List[Dict], review_product: Dict) -> tuple[List[Dict], bool]:
    with st.container(border=True):
        st.subheader("定价")
        saved_settings = review_product.get("pricing_settings", {})
        default_commission_percent = float(saved_settings.get("commission_percent", 15.0))
        pricing_rows = build_pricing_rows(products, review_product.get("pricing_table", []), default_commission_percent)
        pricing_df = pd.DataFrame(pricing_rows).fillna("")

        with st.form("pricing_form", clear_on_submit=False):
            bulk_cost = st.number_input("批量成本价 CNY", min_value=0.0, value=0.0, step=0.1)
            bulk_shipping = st.number_input("批量国内运费 CNY", min_value=0.0, value=0.0, step=0.1)
            bulk_cross_border = st.number_input("批量跨境物流 CNY", min_value=0.0, value=0.0, step=0.1)
            bulk_packaging = st.number_input("批量包装费 CNY", min_value=0.0, value=0.0, step=0.1)
            bulk_other = st.number_input("批量其他成本 CNY", min_value=0.0, value=0.0, step=0.1)
            bulk_commission = st.number_input("批量平台抽佣 %", min_value=0.0, max_value=80.0, value=0.0, step=0.5)
            overwrite_existing = st.checkbox("覆盖已有成本/抽佣", value=True)
            commission_percent = st.number_input("默认平台抽佣 %", min_value=0.0, max_value=80.0, value=default_commission_percent, step=0.5)
            profit_margin_percent = st.number_input("目标毛利率 %", min_value=0.0, max_value=80.0, value=float(saved_settings.get("profit_margin_percent", 25.0)), step=0.5)
            round_step = st.number_input("价格取整到 CNY", min_value=1.0, max_value=100.0, value=float(saved_settings.get("round_step", 1.0)), step=1.0)
            edited_df = st.data_editor(pricing_df, use_container_width=True, hide_index=True, key="maozi_pricing_editor_v3")
            bulk_submit = st.form_submit_button("批量填入成本")
            calc_submit = st.form_submit_button("计算并保存定价")
            save_submit = st.form_submit_button("只保存手动修改")

        calculated_rows = review_product.get("pricing_table") or pricing_rows
        if bulk_submit or calc_submit or save_submit:
            calculated_rows = edited_df.fillna("").to_dict(orient="records")
            if bulk_submit:
                bulk_values = {
                    "成本价CNY": bulk_cost,
                    "国内运费CNY": bulk_shipping,
                    "跨境物流CNY": bulk_cross_border,
                    "包装费CNY": bulk_packaging,
                    "其他成本CNY": bulk_other,
                    "平台抽佣%": bulk_commission,
                }
                for row in calculated_rows:
                    for field, value in bulk_values.items():
                        if value > 0 and (overwrite_existing or not str(row.get(field, "")).strip()):
                            row[field] = value
            if bulk_submit or calc_submit:
                calculated_rows = update_pricing_calculations(calculated_rows, commission_percent, profit_margin_percent, round_step)
            review_product["pricing_table"] = calculated_rows
            review_product["pricing_settings"] = {
                "commission_percent": commission_percent,
                "profit_margin_percent": profit_margin_percent,
                "round_step": round_step,
            }
            st.session_state["review_product"] = review_product
            st.success("定价已保存。")

        st.dataframe(pd.DataFrame(calculated_rows).fillna(""), use_container_width=True, hide_index=True, height=240)
        use_suggested_prices = st.checkbox("生成 Excel 时使用已保存建议售价", value=bool(review_product.get("use_suggested_prices", True)))
        review_product["use_suggested_prices"] = use_suggested_prices
        return calculated_rows, use_suggested_prices


def ai_draft_page() -> None:
    st.header("图片识别生成 Ozon Excel 草稿")
    st.caption("这个页面只生成本地图片和 Excel，不会自动登录 Ozon，也不会自动上架。")

    st.session_state.setdefault("ai_analysis", None)
    st.session_state.setdefault("processed_images", None)
    st.session_state.setdefault("review_product", None)
    st.session_state.setdefault("template_path", "")
    st.session_state.setdefault("ai_offer_id", f"GGYF-{datetime.now().strftime('%H%M%S')}")
    st.session_state.setdefault("category_options", [])
    st.session_state.setdefault("category_description_id", "")
    st.session_state.setdefault("category_type_id", "")
    st.session_state.setdefault("category_type_name", "")
    st.session_state.setdefault("pet_category_template_label", "")
    st.session_state.setdefault("title_template_cn", OZON_TITLE_TEMPLATE_CN)
    st.session_state.setdefault("title_template_ru", OZON_TITLE_TEMPLATE_RU)

    default_image_folder = IMAGE_DIR if list_product_images(IMAGE_DIR) else AI_INPUT_IMAGE_DIR

    with st.container(border=True):
        st.subheader("1. 常用填写")
        offer_id = st.text_input("默认货号（可手动修改）", key="ai_offer_id")
        product_name_cn = st.text_input("商品名称（中文，可自动生成俄文）", value="")
        title_template_cn = st.text_input(
            "推荐商品名称模板",
            value=str(st.session_state.get("title_template_cn", OZON_TITLE_TEMPLATE_CN)),
            help="程序会把这个模板写进 GPT 提示词，并在人工审核区提醒你按这个顺序检查标题。",
        )
        title_template_ru = st.text_input(
            "俄语标题模板",
            value=str(st.session_state.get("title_template_ru", OZON_TITLE_TEMPLATE_RU)),
        )
        st.session_state["title_template_cn"] = title_template_cn.strip() or OZON_TITLE_TEMPLATE_CN
        st.session_state["title_template_ru"] = title_template_ru.strip() or OZON_TITLE_TEMPLATE_RU
        st.info(f"商品名称模板：{st.session_state['title_template_cn']}。注意：标题不写颜色、不写尺码；颜色单独写到颜色字段。")
        st.caption(OZON_TITLE_TEMPLATE_DETAIL)
        with st.expander("查看推荐俄语标题示例", expanded=False):
            st.write(st.session_state["title_template_ru"])
            for example in OZON_TITLE_EXAMPLES_RU:
                st.code(example, language="text")
        model_card_name = st.text_input("型号名称（针对合并为一张商品卡片）", value="", placeholder="手动填写，例如：JSY-救生衣 或 YY-雨衣")
        product_name_ru = st.text_input("商品名称（俄文，自动生成后可修改）", value=translate_product_name_to_ru(product_name_cn))
        clothing_type = st.selectbox("服装类型", CLOTHING_TYPE_OPTIONS, format_func=lambda value: "不选择" if not value else value)
        manual_main_color = st.text_input("手动主色（中文，优先使用）", value="", placeholder="例如：绿色、黑色、红色")
        manual_colors_text = st.text_area("颜色/款式列表", placeholder="每行一个颜色：绿色 或 绿色|зеленый", height=76)
        manual_sizes_text = st.text_area("尺码明细", placeholder="每行一个尺码：S|32-38|25|20-26|价格|重量g|长mm|宽mm|高mm", height=96)
        standard_sizes = st.multiselect("补齐尺码列表", STANDARD_SIZE_OPTIONS, default=DEFAULT_STANDARD_SIZES)
        price = st.number_input("默认价格 CNY", min_value=0.0, value=0.0, step=1.0)
        old_price = st.number_input("折扣前价格 CNY", min_value=0.0, value=0.0, step=1.0)
        stock = st.number_input("库存", min_value=0, value=100, step=1)
        weight = st.number_input("默认重量 g", min_value=0, value=150, step=10)
        product_weight = st.number_input("商品克重 g", min_value=0, value=0, step=1)
        package_dimensions = st.text_input("包装尺寸 mm（长|宽|高）", value="300|220|30")
        depth, width, height = parse_package_dimensions(package_dimensions)

    with st.expander("高级设置：模板、图片文件夹、Ozon 类目", expanded=False):
        uploaded_template = st.file_uploader("选择 Ozon 中文模板 xlsx", type=["xlsx"], key="ai_template")
        template_choices = ozon_template_options()
        template_labels = [str(path) for path in template_choices] or [str(INPUT_DIR / "宠物服装模板.xlsx")]
        selected_template = st.selectbox("选择 Ozon 中文版模板", template_labels, index=0)
        template_text = st.text_input("模板路径", value=selected_template)
        image_folder = st.text_input("商品图片文件夹", value=str(default_image_folder))
        mode = st.selectbox("处理模式", ["单个商品：文件夹内图片都属于同一个商品", "多商品：按文件名前缀分组（后续增强）"])
        model_name = st.text_input("OpenAI 模型（可在 .env 设置 OPENAI_MODEL）", value="")

        pet_category_labels = ["手动填写/不选择"] + [item["label"] for item in PET_CATEGORY_TEMPLATE_OPTIONS]
        current_pet_category_label = str(st.session_state.get("pet_category_template_label", ""))
        pet_category_index = pet_category_labels.index(current_pet_category_label) if current_pet_category_label in pet_category_labels else 0
        selected_pet_category_label = st.selectbox(
            "常用宠物类目（自动填 description_category_id / type_id）",
            pet_category_labels,
            index=pet_category_index,
            format_func=lambda label: "手动填写/不选择" if label == "手动填写/不选择" else pet_category_template_label(PET_CATEGORY_TEMPLATE_BY_LABEL[label]),
        )
        if selected_pet_category_label != "手动填写/不选择":
            selected_pet_category = PET_CATEGORY_TEMPLATE_BY_LABEL[selected_pet_category_label]
            apply_pet_category_template(selected_pet_category)
            st.success(
                f"已选择：{selected_pet_category['label']}，"
                f"description_category_id={selected_pet_category['description_category_id']}，"
                f"type_id={selected_pet_category['type_id']}"
            )

        settings = load_settings()
        if st.button("从 Ozon 拉取类目选项", use_container_width=True):
            category_client = get_client(settings.client_id, settings.api_key, settings.base_url, mock_mode=False)
            category_result = category_client.get_description_category_tree()
            if category_result.get("success"):
                options = flatten_category_tree(category_result.get("response", {}))
                st.session_state["category_options"] = options
                st.success(f"已拉取 {len(options)} 个类目选项。")
            else:
                st.error("拉取类目失败，请检查 Ozon API Key 或稍后重试。")
                st.json(category_result)
        category_keyword = st.text_input("类目关键词筛选", value="宠物服装", placeholder="例如：宠物服装 狗 雨衣")
        category_options = st.session_state.get("category_options", [])
        if category_options:
            render_category_selector(category_options, category_keyword)
        else:
            st.caption("还没有类目选项。可以先手动填写，也可以点击上方按钮从 Ozon 拉取。")

        description_category_id = st.text_input("description_category_id", value=str(st.session_state.get("category_description_id", "")))
        type_id = st.text_input("type_id", value=str(st.session_state.get("category_type_id", "")))
        if st.session_state.get("category_type_name"):
            st.text_input("type_name / 类型名称", value=str(st.session_state.get("category_type_name", "")), disabled=True)

    defaults = {
        "offer_id": offer_id,
        "product_name_cn": product_name_cn,
        "name_ru": product_name_ru,
        "description_category_id": description_category_id,
        "type_id": type_id,
        "type_name": str(st.session_state.get("category_type_name", "")),
        "stock": stock,
        "clothing_type": clothing_type,
        "model": model_card_name.strip(),
        "model_prefix": "",
        "title_template_cn": str(st.session_state.get("title_template_cn", OZON_TITLE_TEMPLATE_CN)),
        "title_template_ru": str(st.session_state.get("title_template_ru", OZON_TITLE_TEMPLATE_RU)),
        "standard_sizes_text": ",".join(standard_sizes),
        "manual_main_color": manual_main_color,
        "price": price if price > 0 else "",
        "old_price": old_price if old_price > 0 else "",
        "weight": weight,
        "product_weight": product_weight if product_weight > 0 else "",
        "depth": depth,
        "width": width,
        "height": height,
    }

    images = list_product_images(image_folder)
    if images:
        st.success(f"已找到 {len(images)} 张图片。")
        with st.expander("查看图片预览", expanded=False):
            st.image([str(path) for path in images[:8]], caption=[path.name for path in images[:8]], width=140)
    else:
        st.warning("还没有找到图片，请把商品图片放进上面的文件夹，或改成你的图片目录。")

    st.subheader("2. 识别方式")
    gpt_prompt = build_gpt_plus_prompt(
        [path.name for path in images],
        product_name_cn=product_name_cn,
        colors_text=manual_colors_text or manual_main_color,
        sizes_text=manual_sizes_text,
        title_template_cn=str(st.session_state.get("title_template_cn", OZON_TITLE_TEMPLATE_CN)),
        title_template_ru=str(st.session_state.get("title_template_ru", OZON_TITLE_TEMPLATE_RU)),
    )
    st.link_button("打开 ChatGPT Plus", "https://chatgpt.com/")
    render_copy_prompt(gpt_prompt)
    with st.expander("查看完整 GPT 提示词", expanded=False):
        st.text_area("提示词内容", value=gpt_prompt, height=220)
    uploaded_gpt_json = st.file_uploader("上传 GPT 返回 txt/json", type=["txt", "json"], key="gpt_json_file")
    uploaded_gpt_text = uploaded_gpt_json.getvalue().decode("utf-8", errors="ignore") if uploaded_gpt_json else ""
    pasted_gpt_json = st.text_area("粘贴 ChatGPT 返回的 JSON", value=uploaded_gpt_text, height=140)
    import_gpt_json = st.button("导入 GPT Plus 识别结果", disabled=not images or not pasted_gpt_json.strip())

    has_custom_gpt = custom_gpt_configured()
    col_run, col_mock = st.columns(2)
    with col_run:
        run_ai = st.button("使用自定义 GPT 接口识别", type="primary", disabled=(not images or not has_custom_gpt))
    with col_mock:
        run_mock = st.button("不用 OpenAI，手动建商品草稿", disabled=not images)
    if mode.startswith("多商品"):
        st.info("当前先稳定支持单个商品。")

    if run_ai or run_mock or import_gpt_json:
        try:
            with st.spinner("正在识别图片并整理商品资料..."):
                if import_gpt_json:
                    analysis = extract_json_object(pasted_gpt_json)
                elif run_mock:
                    analysis = mock_product_analysis([str(path) for path in images])
                else:
                    analysis = analyze_product_images([str(path) for path in images], model=model_name.strip() or None)
                json_path = save_json(analysis, OUTPUT_DIR / "product_json" / f"{safe_offer_id(offer_id)}.json")
                processed = process_product_images(images, AI_OUTPUT_IMAGE_DIR, offer_id, analysis.get("image_roles", {}))
                review_product = build_review_product(analysis, processed, defaults)
                if manual_colors_text.strip():
                    review_product["color_variants_text"] = manual_colors_text.strip()
                elif manual_main_color.strip() and not format_color_variants(analysis).strip():
                    normalized_color = normalize_color_cn(manual_main_color)
                    review_product["color_variants_text"] = normalized_color
                    review_product["color"] = normalized_color
                    review_product["color_name_ru"] = normalized_color
                if manual_sizes_text.strip():
                    review_product["size_variants_text"] = manual_sizes_text.strip()
                st.session_state["ai_analysis"] = analysis
                st.session_state["processed_images"] = processed
                st.session_state["review_product"] = review_product
                st.success(f"商品草稿已保存为 JSON：{json_path}")
        except Exception as exc:
            st.error(str(exc))

    analysis = st.session_state.get("ai_analysis")
    processed = st.session_state.get("processed_images")
    review_product = st.session_state.get("review_product")
    if not (analysis and processed and review_product):
        return

    st.subheader("识别结果预览")
    st.write("商品类型：", analysis.get("product_type", ""))
    st.write("俄文类型：", analysis.get("product_type_ru", ""))
    st.write("置信度：", analysis.get("confidence", ""))
    if analysis.get("manual_review_notes"):
        st.warning("；".join(analysis.get("manual_review_notes", [])))

    st.subheader("人工审核和修改")
    with st.form("review_form"):
        edited_offer_id = st.text_input("货号*", value=str(review_product.get("offer_id", "")))
        edited_name_cn = st.text_input("商品名称（中文）", value=str(review_product.get("product_name_cn", "")))
        edited_model = st.text_input("型号名称（针对合并为一张商品卡片）", value=str(review_product.get("model", "")))
        review_title_template_cn = str(review_product.get("title_template_cn") or st.session_state.get("title_template_cn", OZON_TITLE_TEMPLATE_CN))
        review_title_template_ru = str(review_product.get("title_template_ru") or st.session_state.get("title_template_ru", OZON_TITLE_TEMPLATE_RU))
        st.caption(f"标题模板：{review_title_template_cn}；颜色和尺码不要写进商品名称。")
        with st.expander("标题模板说明", expanded=False):
            st.write("中文模板：", review_title_template_cn)
            st.write("俄语模板：", review_title_template_ru)
            st.write(OZON_TITLE_TEMPLATE_DETAIL)
        edited_name = st.text_input("商品名称（俄文，可手动覆盖）", value=str(review_product.get("name_ru", "")))
        review_pet_category_labels = ["手动填写/不选择"] + [item["label"] for item in PET_CATEGORY_TEMPLATE_OPTIONS]
        review_current_type_name = str(review_product.get("type_name", "") or st.session_state.get("category_type_name", ""))
        review_category_index = review_pet_category_labels.index(review_current_type_name) if review_current_type_name in review_pet_category_labels else 0
        edited_pet_category_label = st.selectbox(
            "常用宠物类目（自动填 ID）",
            review_pet_category_labels,
            index=review_category_index,
            format_func=lambda label: "手动填写/不选择" if label == "手动填写/不选择" else pet_category_template_label(PET_CATEGORY_TEMPLATE_BY_LABEL[label]),
        )
        if edited_pet_category_label != "手动填写/不选择":
            edited_pet_category = PET_CATEGORY_TEMPLATE_BY_LABEL[edited_pet_category_label]
            edited_description_category_id_default = edited_pet_category["description_category_id"]
            edited_type_id_default = edited_pet_category["type_id"]
            edited_type_name_default = edited_pet_category["type_name"]
        else:
            edited_description_category_id_default = str(review_product.get("description_category_id", ""))
            edited_type_id_default = str(review_product.get("type_id", ""))
            edited_type_name_default = str(review_product.get("type_name", ""))
        edited_description_category_id = st.text_input("description_category_id", value=str(edited_description_category_id_default))
        edited_type_id = st.text_input("type_id", value=str(edited_type_id_default))
        edited_type_name = st.text_input("type_name / 类型名称", value=str(edited_type_name_default))
        edited_stock = st.number_input("库存", min_value=0, value=int(number_value(review_product.get("stock", 100), 100)), step=1)
        edited_price = st.text_input("价格，CNY*", value=str(review_product.get("price", "")))
        edited_old_price = st.text_input("折扣前价格，CNY", value=str(review_product.get("old_price", "")))
        current_clothing_type = str(review_product.get("clothing_type", ""))
        clothing_index = CLOTHING_TYPE_OPTIONS.index(current_clothing_type) if current_clothing_type in CLOTHING_TYPE_OPTIONS else 0
        edited_clothing_type = st.selectbox("服装类型", CLOTHING_TYPE_OPTIONS, index=clothing_index, format_func=lambda value: "不选择" if not value else value)
        edited_color = st.text_input("商品主色", value=str(review_product.get("color", "")))
        edited_color_name = st.text_input("颜色名称（默认同主色）", value=str(review_product.get("color_name_ru", "")))
        edited_colors_text = st.text_area("颜色/款式列表", value=str(review_product.get("color_variants_text", "")), height=76)
        edited_sizes_text = st.text_area("尺码明细", value=str(review_product.get("size_variants_text", "")), height=96)
        edited_standard_sizes = st.multiselect(
            "补齐尺码列表",
            STANDARD_SIZE_OPTIONS,
            default=parse_standard_sizes(review_product.get("standard_sizes_text", "")) or DEFAULT_STANDARD_SIZES,
        )
        edited_desc = st.text_area("简介（俄文）", value=str(review_product.get("description_ru", "")), height=110)
        edited_tags = st.text_area("俄文标签", value=" ".join(review_product.get("tags_ru", [])), height=96)
        target_pet_value = str(review_product.get("target_pet") or target_pet_cn_from_context(edited_name_cn, analysis.get("target_audience", "")))
        target_pet_index = TARGET_PET_CN_OPTIONS.index(target_pet_value) if target_pet_value in TARGET_PET_CN_OPTIONS else 0
        edited_target_pet = st.selectbox("专为", TARGET_PET_CN_OPTIONS, index=target_pet_index)
        edited_weight = st.text_input("毛重，克*", value=str(review_product.get("weight", "")))
        edited_product_weight = st.text_input("商品克重，克", value=str(review_product.get("product_weight", "")))
        edited_dimensions = st.text_input(
            "包装尺寸 mm（长|宽|高）",
            value="|".join(str(review_product.get(field, "")) for field in ["depth", "width", "height"]),
        )
        edited_depth, edited_width, edited_height = parse_package_dimensions(
            edited_dimensions,
            default_depth=int(number_value(review_product.get("depth", 300), 300)),
            default_width=int(number_value(review_product.get("width", 220), 220)),
            default_height=int(number_value(review_product.get("height", 30), 30)),
        )
        edited_material = st.text_input("材料", value=str(review_product.get("material", "")))
        confirmed = st.form_submit_button("确认这些资料")

    if confirmed:
        title_source = {
            **review_product,
            "product_name_cn": edited_name_cn,
            "name_ru": edited_name,
            "color": normalize_color_cn(edited_color),
            "clothing_type": edited_clothing_type,
            "style": analysis.get("style", ""),
            "season": analysis.get("season", ""),
            "target_pet": edited_target_pet,
        }
        final_name_ru = generate_ozon_title_ru(title_source, analysis)
        if edited_name.strip() and edited_name.strip() != str(review_product.get("name_ru", "")).strip():
            final_name_ru = clean_ozon_title_ru(edited_name)

        review_product.update(
            {
                "offer_id": safe_offer_id(edited_offer_id),
                "product_name_cn": edited_name_cn,
                "name_ru": final_name_ru,
                "description_ru": edited_desc or make_description_ru(final_name_ru),
                "description_category_id": edited_description_category_id.strip(),
                "type_id": edited_type_id.strip(),
                "type_name": edited_type_name.strip(),
                "stock": edited_stock,
                "tags_ru": ensure_search_tags_ru(
                    {**review_product, "product_name_cn": edited_name_cn, "name_ru": final_name_ru, "clothing_type": edited_clothing_type},
                    [item for item in re.split(r"[\s,，]+", edited_tags) if item.strip()],
                ),
                "price": edited_price,
                "old_price": edited_old_price,
                "weight": edited_weight,
                "product_weight": edited_product_weight,
                "depth": edited_depth,
                "width": edited_width,
                "height": edited_height,
                "material": edited_material,
                "color": normalize_color_cn(edited_color),
                "color_name_ru": guess_color_name_ru(edited_color, edited_color_name),
                "clothing_type": edited_clothing_type,
                "model_prefix": "",
                "title_template_cn": review_title_template_cn,
                "title_template_ru": review_title_template_ru,
                "color_variants_text": edited_colors_text,
                "size_variants_text": edited_sizes_text,
                "standard_sizes_text": ",".join(edited_standard_sizes),
                "model": edited_model.strip(),
                "animal_gender": "男女两用的",
                "target_pet": edited_target_pet,
                "unit_product_quantity": 1,
                "items_in_product": 1,
                "brand_country": "中国",
                "origin_country": "中国",
            }
        )
        product_json_dir = OUTPUT_DIR / "product_json"
        product_json_dir.mkdir(parents=True, exist_ok=True)
        review_product_path = save_json(review_product, product_json_dir / f"{review_product['offer_id']}_reviewed.json")
        st.session_state["review_product"] = review_product
        st.session_state["review_product_json_path"] = str(review_product_path)
        st.success("已确认，下一步可以生成 Excel。")

    st.subheader("生成 Excel")
    preview_products = build_variant_products(st.session_state["review_product"])
    pricing_rows, use_suggested_prices = render_pricing_tool(preview_products, st.session_state["review_product"])
    if use_suggested_prices:
        preview_products = apply_pricing_to_products(preview_products, pricing_rows)
    preview_cols = [
        "offer_id",
        "model",
        "color",
        "pet_size",
        "target_pet",
        "price",
        "old_price",
        "stock",
        "weight",
        "product_weight",
        "description_category_id",
        "type_id",
    ]
    preview_df = pd.DataFrame(preview_products)
    st.dataframe(preview_df[[col for col in preview_cols if col in preview_df.columns]], use_container_width=True, height=280)
    st.caption(f"将生成 {len(preview_products)} 个商品行。")

    template_path = save_uploaded_excel(uploaded_template) if uploaded_template else Path(template_text)
    output_name = f"Ozon_AI_已填写_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    output_path = AI_OUTPUT_EXCEL_DIR / output_name
    if st.button("生成已填写 Excel", type="primary"):
        if not template_path.exists():
            st.error(f"找不到模板：{template_path}")
        else:
            try:
                final_product = st.session_state["review_product"]
                final_products = build_variant_products(final_product)
                if use_suggested_prices:
                    final_products = apply_pricing_to_products(final_products, final_product.get("pricing_table", []))
                result_path = fill_ozon_template(template_path, final_products, output_path)
                report = {
                    "created_at": datetime.now().isoformat(timespec="seconds"),
                    "excel": str(result_path),
                    "images": processed,
                    "products": final_products,
                    "product_count": len(final_products),
                    "note": "本地生成文件，请人工核对后再用于 Ozon。",
                }
                report_path = save_json(report, OUTPUT_DIR / "处理报告.json")
                st.session_state["last_generated_excel_path"] = str(result_path)
                st.session_state["last_generated_image_root"] = str(processed.get("folder", AI_OUTPUT_IMAGE_DIR))
                st.session_state["auto_load_generated_excel"] = True
                st.success(f"Excel 已生成：{result_path}，共 {len(final_products)} 个商品行。")
                st.info(f"处理报告：{report_path}")
                st.download_button(
                    "下载 Excel",
                    data=result_path.read_bytes(),
                    file_name=result_path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as exc:
                st.error(f"生成 Excel 失败：{exc}")


def api_upload_page() -> None:
    st.header("Ozon API 上传工具")
    settings = load_settings()
    st.session_state.setdefault("products", [])
    st.session_state.setdefault("validation_results", [])
    with st.container(border=True):
        dry_run = st.toggle("dry-run 预检查", value=True)
        mock_mode = st.toggle("mock 模式（不真实上传）", value=settings.mock_mode)
        loose_draft_mode = st.toggle("宽松草稿模式", value=True)
        mode = st.selectbox("上传模式", ["只校验", "上传商品", "上传图片", "更新价格", "更新库存", "全流程"])
        uploaded_excel = st.file_uploader("上传商品 Excel", type=["xlsx"], key="api_excel")
        generated_excel = st.session_state.get("last_generated_excel_path", "")
        default_image_root = Path(st.session_state.get("last_generated_image_root", IMAGE_DIR))
        image_root = Path(st.text_input("图片根目录", value=str(default_image_root)))
        excel_to_load = None
        if uploaded_excel:
            excel_to_load = save_uploaded_excel(uploaded_excel)
        elif generated_excel and Path(generated_excel).exists() and st.button("载入刚生成的 Excel", use_container_width=True):
            excel_to_load = Path(generated_excel)
        if excel_to_load:
            loaded = load_excel_into_api_session(excel_to_load, image_root, dry_run, mode, loose_draft_mode)
            st.success(f"已载入 {loaded['row_count']} 行商品。")

    products = st.session_state.get("products", [])
    if products:
        refresh_api_validation(dry_run, mode, loose_draft_mode)
        validation_results = st.session_state.get("validation_results", [])
        visible_cols = ["offer_id", "name_ru", "color", "size", "price", "stock", "description_category_id", "type_id"]
        preview_df = pd.DataFrame(products).fillna("")
        st.dataframe(preview_df[[col for col in visible_cols if col in preview_df.columns]], use_container_width=True, height=280)
        validation_df = pd.DataFrame(validation_results)
        if not validation_df.empty:
            st.dataframe(validation_df, use_container_width=True, height=220)
        errors = [item for item in validation_results if item.get("status") == "错误"]
        client = get_client(settings.client_id, settings.api_key, settings.base_url, mock_mode=mock_mode)
        if st.button("执行当前上传模式", disabled=dry_run or bool(errors), use_container_width=True):
            result = run_api_step(client, mode, products)
            st.json(result)
    else:
        st.info("请先生成或载入商品 Excel。")


def main() -> None:
st.set_page_config(
    page_title="Ozon 商品 AI 半自动生成助手",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================
# Ozon API 输入窗口
# =========================
ENV_PATH = BASE_DIR / ".env"
load_dotenv(ENV_PATH)

with st.sidebar.expander("🔑 Ozon API 设置", expanded=True):
    ozon_client_id_input = st.text_input(
        "Client-Id",
        value=os.getenv("OZON_CLIENT_ID", ""),
        help="填写 Ozon Seller 后台的 Client-Id",
    )
    ozon_api_key_input = st.text_input(
        "Api-Key",
        value=os.getenv("OZON_API_KEY", ""),
        type="password",
        help="填写 Ozon Seller 后台的 Api-Key",
    )

    if st.button("保存 Ozon API"):
        ENV_PATH.touch(exist_ok=True)
        set_key(str(ENV_PATH), "OZON_CLIENT_ID", ozon_client_id_input.strip())
        set_key(str(ENV_PATH), "OZON_API_KEY", ozon_api_key_input.strip())
        os.environ["OZON_CLIENT_ID"] = ozon_client_id_input.strip()
        os.environ["OZON_API_KEY"] = ozon_api_key_input.strip()
        st.success("Ozon API 已保存到本地 .env")
    st.sidebar.header("页面导航")
    page = st.sidebar.radio("选择功能", ["AI 生成 Excel 草稿", "Ozon API 上传"])
    st.sidebar.caption("在左侧切换页面，不需要回到顶部。")
    st.title("Ozon 商品 AI 半自动生成助手")
    if page == "AI 生成 Excel 草稿":
        ai_draft_page()
    else:
        api_upload_page()


if __name__ == "__main__":
    main()
