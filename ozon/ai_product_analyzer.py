from __future__ import annotations

import base64
import json
import mimetypes
import os
import re
from pathlib import Path
from typing import Any, Dict, List

from dotenv import load_dotenv

from config import BASE_DIR, LOG_DIR


PRODUCT_SCHEMA: Dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "product_type": {"type": "string"},
        "product_type_ru": {"type": "string"},
        "material": {"type": "string"},
        "color": {"type": "string"},
        "color_name_ru": {"type": "string"},
        "color_variants": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "color": {"type": "string"},
                    "color_name_ru": {"type": "string"},
                    "note": {"type": "string"},
                },
                "required": ["color", "color_name_ru", "note"],
            },
        },
        "sizes": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "size": {"type": "string"},
                    "chest_cm": {"type": "string"},
                    "back_length_cm": {"type": "string"},
                    "neck_cm": {"type": "string"},
                    "price": {"type": "string"},
                    "weight": {"type": "string"},
                    "depth": {"type": "string"},
                    "width": {"type": "string"},
                    "height": {"type": "string"},
                    "note": {"type": "string"},
                },
                "required": [
                    "size",
                    "chest_cm",
                    "back_length_cm",
                    "neck_cm",
                    "price",
                    "weight",
                    "depth",
                    "width",
                    "height",
                    "note",
                ],
            },
        },
        "target_audience": {"type": "string"},
        "style": {"type": "string"},
        "season": {"type": "string"},
        "main_features": {"type": "array", "items": {"type": "string"}},
        "name_ru": {"type": "string"},
        "description_ru": {"type": "string"},
        "bullets_ru": {"type": "array", "items": {"type": "string"}},
        "tags_ru": {"type": "array", "items": {"type": "string"}},
        "search_keywords_ru": {"type": "array", "items": {"type": "string"}},
        "recommended_category_note": {"type": "string"},
        "image_roles": {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "primary_image": {"type": "string"},
                "additional_images": {"type": "array", "items": {"type": "string"}},
                "size_chart": {"type": "string"},
                "detail_images": {"type": "array", "items": {"type": "string"}},
            },
            "required": ["primary_image", "additional_images", "size_chart", "detail_images"],
        },
        "confidence": {"type": "number"},
        "need_manual_review": {"type": "boolean"},
        "manual_review_notes": {"type": "array", "items": {"type": "string"}},
    },
    "required": [
        "product_type",
        "product_type_ru",
        "material",
        "color",
        "color_name_ru",
        "color_variants",
        "sizes",
        "target_audience",
        "style",
        "season",
        "main_features",
        "name_ru",
        "description_ru",
        "bullets_ru",
        "tags_ru",
        "search_keywords_ru",
        "recommended_category_note",
        "image_roles",
        "confidence",
        "need_manual_review",
        "manual_review_notes",
    ],
}


def _image_to_data_url(path: Path) -> str:
    mime_type = mimetypes.guess_type(path.name)[0] or "image/png"
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime_type};base64,{encoded}"


def _extract_json(text: str) -> Dict[str, Any]:
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?", "", cleaned, flags=re.IGNORECASE).strip()
        cleaned = re.sub(r"```$", "", cleaned).strip()
    match = re.search(r"\{.*\}", cleaned, flags=re.DOTALL)
    if match:
        cleaned = match.group(0)
    return json.loads(cleaned)


def _save_failed_response(raw_text: str) -> Path:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    output = LOG_DIR / "ai_failed_response.txt"
    output.write_text(raw_text, encoding="utf-8")
    return output


def _openai_error_message(exc: Exception) -> str:
    text = str(exc)
    lowered = text.lower()
    if "insufficient_quota" in lowered or "exceeded your current quota" in lowered:
        return "OpenAI 额度不足或账号计费不可用：请检查 API Key 对应账号的余额/套餐/账单设置，或换一个有额度的 Key。"
    if "invalid_api_key" in lowered or "incorrect api key" in lowered:
        return "OpenAI API Key 无效：请检查 .env 里的 OPENAI_API_KEY 是否复制完整，前后不要有空格。"
    if "model_not_found" in lowered or "does not exist" in lowered:
        return "OpenAI 模型不可用：请检查 OPENAI_MODEL，或换成你的账号支持的视觉模型。"
    if "rate_limit" in lowered or "error code: 429" in lowered:
        return "OpenAI 请求过快或达到限额：请稍后再试，或检查账号限额。"
    return f"AI 调用失败：{text}"


def _prompt(file_names: List[str]) -> str:
    return (
        "你是 Ozon 俄罗斯电商宠物服装商品资料整理助手。请只根据图片中能确认的信息提取商品资料，"
        "不要编造看不见的信息。俄文标题、简介和标签要自然，适合 Ozon 搜索和俄罗斯买家理解。"
        "name_ru 必须是 1 个最佳俄语商品标题，像 Ozon 真实商品标题。固定标题模板：商品类型 + 款式/功能 + 适用对象 + 使用场景/核心卖点；俄语顺序参考：Тип товара + фасон/функция + для кого + сценарий/ключевое преимущество。"
        "宠物服装标题优先使用 одежда、костюм、куртка、жилет、плащ、комбинезон、худи 等自然品类词；适用对象用 для кошек、для собак、для щенков、для маленьких собак、для питомцев。"
        "不要在 name_ru 里写颜色、尺码、SKU、成本、库存、物流词、爆款、热卖、跨境专用、лучший、супер、хит、распродажа、表情符号或无关关键词。"
        "颜色必须只写到 color 和 color_name_ru 字段；尺码只写入 sizes，不放进标题。标题长度控制在 60 到 120 个俄文字符左右，最长不超过 255 个字符。"
        "不要使用夸大词、医疗功效词、官方正品、最便宜、品牌侵权词。"
        "如果材质、重量、包装尺寸、尺码等无法从图片确认，请留空并在 manual_review_notes 说明。"
        "如果图片或尺码表里能看到多个颜色/款式，请写入 color_variants；如果只能确认一个颜色，color_variants 可以为空。"
        "如果能看到多个尺码，请写入 sizes，并尽量提取每个尺码对应的胸围 chest_cm、背长 back_length_cm、颈围 neck_cm。"
        "如果同一款商品有多个颜色，请把 sizes 当作整组商品的统一尺码池，不要让某个颜色漏掉 XS、S 等尺码。"
        "胸围/颈围如果是范围，按 32-38 这种格式返回；无法确认价格、重量、包装长宽高时这些字段留空，不要估算。"
        "请识别哪张是主图、附图、尺码图、详情图。文件名列表："
        + json.dumps(file_names, ensure_ascii=False)
        + "。必须返回一个 JSON 对象，字段严格匹配给定结构。tags_ru 需要 15 到 20 个，以 # 开头；标签要符合俄罗斯买家的搜索习惯，优先使用自然俄语搜索短语，覆盖品类、宠物对象、用途、季节、场景和常见同义词。"
    )


def _custom_gpt_url() -> str:
    return os.getenv("CUSTOM_GPT_URL", "").strip() or os.getenv("GPT_PROGRAM_URL", "").strip()


def _analyze_with_custom_gpt(image_paths: List[str], endpoint: str, model: str | None = None) -> Dict[str, Any]:
    import requests

    paths = [Path(item) for item in image_paths]
    headers = {"Content-Type": "application/json"}
    custom_key = os.getenv("CUSTOM_GPT_API_KEY", "").strip() or os.getenv("GPT_PROGRAM_API_KEY", "").strip()
    if custom_key:
        headers["Authorization"] = f"Bearer {custom_key}"

    payload: Dict[str, Any] = {
        "model": model or os.getenv("OPENAI_MODEL", "gpt-4.1-mini"),
        "prompt": _prompt([p.name for p in paths]),
        "images": [{"name": path.name, "data_url": _image_to_data_url(path)} for path in paths],
        "schema": PRODUCT_SCHEMA,
    }
    response = requests.post(endpoint, headers=headers, json=payload, timeout=180)
    raw_text = response.text
    if response.status_code >= 400:
        failed_path = _save_failed_response(raw_text)
        raise RuntimeError(f"自定义 GPT 程序返回错误 HTTP {response.status_code}，原始内容已保存：{failed_path}")

    try:
        data = response.json()
    except ValueError:
        return _extract_json(raw_text)

    if isinstance(data, dict):
        for key in ["result", "analysis", "data", "json"]:
            value = data.get(key)
            if isinstance(value, dict):
                return value
            if isinstance(value, str):
                return _extract_json(value)
        text = data.get("text") or data.get("output_text") or data.get("content")
        if isinstance(text, str):
            return _extract_json(text)
        return data
    if isinstance(data, str):
        return _extract_json(data)
    raise RuntimeError("自定义 GPT 程序返回的内容不是 JSON 对象。")


def analyze_product_images(image_paths: List[str], model: str | None = None) -> Dict[str, Any]:
    load_dotenv(BASE_DIR / ".env")
    custom_endpoint = _custom_gpt_url()
    if custom_endpoint:
        try:
            return _analyze_with_custom_gpt(image_paths, custom_endpoint, model=model)
        except Exception as exc:
            failed_path = _save_failed_response(str(exc))
            raise RuntimeError(f"自定义 GPT 程序调用失败：{exc} 原始错误已保存：{failed_path}") from exc

    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("未检测到 OPENAI_API_KEY，也没有配置 CUSTOM_GPT_URL。请先配置 API Key、自定义 GPT 程序地址，或使用手动建商品草稿。")

    from openai import APIError, OpenAI

    paths = [Path(item) for item in image_paths]
    model_name = model or os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
    client = OpenAI(api_key=api_key)

    content: List[Dict[str, Any]] = [{"type": "input_text", "text": _prompt([p.name for p in paths])}]
    for path in paths:
        content.append({"type": "input_image", "image_url": _image_to_data_url(path)})

    last_raw = ""
    for attempt in range(2):
        try:
            kwargs: Dict[str, Any] = {
                "model": model_name,
                "input": [{"role": "user", "content": content}],
            }
            if attempt == 0:
                kwargs["text"] = {
                    "format": {
                        "type": "json_schema",
                        "name": "ozon_pet_product_analysis",
                        "schema": PRODUCT_SCHEMA,
                        "strict": True,
                    }
                }
            response = client.responses.create(**kwargs)
            last_raw = getattr(response, "output_text", "") or str(response)
            return _extract_json(last_raw)
        except APIError as exc:
            failed_path = _save_failed_response(str(exc))
            raise RuntimeError(f"{_openai_error_message(exc)} 原始错误已保存：{failed_path}") from exc
        except (json.JSONDecodeError, ValueError) as exc:
            if attempt == 1:
                failed_path = _save_failed_response(last_raw or str(exc))
                raise RuntimeError(f"AI 返回内容无法解析为 JSON，原始内容已保存：{failed_path}") from exc
        except Exception as exc:
            last_raw = last_raw or str(exc)
            if attempt == 1:
                failed_path = _save_failed_response(last_raw)
                raise RuntimeError(f"{_openai_error_message(exc)} 原始错误已保存：{failed_path}") from exc

    raise RuntimeError("AI 识别失败，请稍后重试。")


def mock_product_analysis(image_paths: List[str]) -> Dict[str, Any]:
    names = [Path(path).name for path in image_paths]
    primary = names[0] if names else ""
    size_chart = next((name for name in names if any(k in name.lower() for k in ["size", "chart", "尺寸", "尺码"])), "")
    details = [name for name in names if name not in {primary, size_chart}]
    return {
        "product_type": "宠物雨衣",
        "product_type_ru": "дождевик для собак",
        "material": "",
        "color": "",
        "color_name_ru": "",
        "color_variants": [],
        "sizes": [
            {
                "size": "XS",
                "chest_cm": "",
                "back_length_cm": "",
                "neck_cm": "",
                "price": "",
                "weight": "",
                "depth": "",
                "width": "",
                "height": "",
                "note": "模拟尺码",
            },
            {
                "size": "S",
                "chest_cm": "",
                "back_length_cm": "",
                "neck_cm": "",
                "price": "",
                "weight": "",
                "depth": "",
                "width": "",
                "height": "",
                "note": "模拟尺码",
            },
            {
                "size": "M",
                "chest_cm": "",
                "back_length_cm": "",
                "neck_cm": "",
                "price": "",
                "weight": "",
                "depth": "",
                "width": "",
                "height": "",
                "note": "模拟尺码",
            },
            {
                "size": "L",
                "chest_cm": "",
                "back_length_cm": "",
                "neck_cm": "",
                "price": "",
                "weight": "",
                "depth": "",
                "width": "",
                "height": "",
                "note": "模拟尺码",
            },
            {
                "size": "XL",
                "chest_cm": "",
                "back_length_cm": "",
                "neck_cm": "",
                "price": "",
                "weight": "",
                "depth": "",
                "width": "",
                "height": "",
                "note": "模拟尺码",
            },
            {
                "size": "XXL",
                "chest_cm": "",
                "back_length_cm": "",
                "neck_cm": "",
                "price": "",
                "weight": "",
                "depth": "",
                "width": "",
                "height": "",
                "note": "模拟尺码",
            },
        ],
        "target_audience": "狗",
        "style": "户外防雨",
        "season": "春秋雨天",
        "main_features": ["防雨", "适合户外散步", "宠物服装"],
        "name_ru": "Дождевик для собак для прогулок в дождливую погоду",
        "description_ru": (
            "Лёгкий дождевик для собак подходит для прогулок в дождливую погоду. "
            "Модель помогает защитить шерсть от влаги и грязи. Перед покупкой проверьте размер по таблице."
        ),
        "bullets_ru": ["Для прогулок в дождь", "Лёгкая посадка", "Подходит для собак разных размеров"],
        "tags_ru": [
            "#дождевик_для_собак",
            "#одежда_для_собак",
            "#защита_от_дождя",
            "#для_прогулок",
            "#товары_для_питомцев",
        ],
        "search_keywords_ru": ["дождевик для собак", "одежда для собак", "плащ для собаки"],
        "recommended_category_note": "模拟数据：请人工确认 Ozon 类目 ID 和 type_id。",
        "image_roles": {
            "primary_image": primary,
            "additional_images": details,
            "size_chart": size_chart,
            "detail_images": details,
        },
        "confidence": 0.5,
        "need_manual_review": True,
        "manual_review_notes": ["这是模拟 AI 数据，仅用于测试 Excel 写入和图片整理。"],
    }
