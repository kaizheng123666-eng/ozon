from __future__ import annotations

import json
from collections import Counter
from typing import Dict, List


POSITIVE_NUMBER_FIELDS = ["price", "weight", "depth", "width", "height"]
NON_NEGATIVE_NUMBER_FIELDS = ["stock"]
SOFT_REQUIRED_FIELDS = ["description_ru", "description_category_id", "type_id"]


def _is_blank(value) -> bool:
    return value is None or str(value).strip() == ""


def _to_float(value):
    if _is_blank(value):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def validate_products(
    products: List[Dict],
    require_public_image_urls: bool = False,
    require_image_fields: bool = True,
    loose_draft_mode: bool = False,
) -> List[Dict]:
    offer_counts = Counter(str(p.get("offer_id", "")).strip() for p in products)
    results = []

    for index, product in enumerate(products, start=2):
        errors: List[str] = []
        warnings: List[str] = []
        offer_id = str(product.get("offer_id", "")).strip()

        if not offer_id:
            errors.append("offer_id 不能为空")
        elif offer_counts[offer_id] > 1:
            errors.append(f"offer_id 重复：{offer_id}")

        if _is_blank(product.get("name_ru")):
            errors.append("name_ru 不能为空")

        for field in SOFT_REQUIRED_FIELDS:
            if _is_blank(product.get(field)):
                message = f"{field} 为空"
                if loose_draft_mode:
                    warnings.append(f"{message}；宽松草稿模式下不阻断，本地会尝试提交给 Ozon")
                else:
                    errors.append(f"{field} 不能为空")

        for field in POSITIVE_NUMBER_FIELDS:
            value = _to_float(product.get(field))
            if value is None or value <= 0:
                message = f"{field} 必须大于 0"
                if loose_draft_mode:
                    warnings.append(f"{message}；宽松草稿模式下不阻断，本地会尝试提交给 Ozon")
                else:
                    errors.append(message)

        for field in NON_NEGATIVE_NUMBER_FIELDS:
            value = _to_float(product.get(field))
            if value is None or value < 0:
                message = f"{field} 必须大于等于 0"
                if loose_draft_mode:
                    warnings.append(f"{message}；宽松草稿模式下上传时会按 0 处理")
                else:
                    errors.append(message)

        if _is_blank(product.get("currency_code")):
            warnings.append("currency_code 为空，上传时将默认 RUB")
        if _is_blank(product.get("vat")):
            warnings.append("vat 为空，上传时将默认 0")

        attributes_json = product.get("attributes_json", "")
        try:
            parsed_attributes = json.loads(attributes_json or "{}")
            if not isinstance(parsed_attributes, (dict, list)):
                message = "attributes_json 必须是 JSON 对象或数组"
                if loose_draft_mode:
                    warnings.append(f"{message}；宽松草稿模式下上传时会按空属性处理")
                else:
                    errors.append(message)
        except json.JSONDecodeError as exc:
            message = f"attributes_json 不是合法 JSON：{exc}"
            if loose_draft_mode:
                warnings.append(f"{message}；宽松草稿模式下上传时会按空属性处理")
            else:
                errors.append(message)

        primary = product.get("_primary_image_resolved", {})
        if not primary.get("exists"):
            message = primary.get("error") or "主图不存在"
            if require_image_fields:
                errors.append(message)
            else:
                warnings.append(f"{message}；真实上传图片前需要填写公网图片链接或手动上传图片")

        for item in product.get("_images_resolved", []):
            if not item.get("exists"):
                message = item.get("error") or "附图不存在"
                if require_image_fields:
                    errors.append(message)
                else:
                    warnings.append(f"{message}；真实上传图片前需要填写公网图片链接或手动上传图片")

        if require_public_image_urls and not product.get("_api_image_urls"):
            errors.append("Ozon API 需要可访问图片 URL；当前只有本地图片路径，不能真实上传图片")
        elif product.get("_local_image_paths"):
            warnings.append("当前图片为本地文件，真实上传图片前需要图床/对象存储 URL")

        status = "通过"
        if warnings:
            status = "警告"
        if errors:
            status = "错误"

        results.append(
            {
                "excel_row": index,
                "offer_id": offer_id,
                "status": status,
                "success": not errors,
                "errors": errors,
                "warnings": warnings,
                "error_message": "；".join(errors),
                "warning_message": "；".join(warnings),
            }
        )

    return results


def split_valid_products(products: List[Dict], validation_results: List[Dict]) -> List[Dict]:
    valid_offer_ids = {item["offer_id"] for item in validation_results if item["success"]}
    return [product for product in products if str(product.get("offer_id", "")).strip() in valid_offer_ids]
