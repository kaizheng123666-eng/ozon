from __future__ import annotations

from datetime import datetime
import json
import logging
from pathlib import Path
from typing import Any, Dict, List

import requests

from config import LOG_DIR, mask_secret


def setup_logger() -> logging.Logger:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("ozon_auto_listing")
    logger.setLevel(logging.INFO)
    if logger.handlers:
        return logger

    log_path = LOG_DIR / f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    handler = logging.FileHandler(log_path, encoding="utf-8")
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.info("日志文件：%s", log_path)
    return logger


class OzonClient:
    def __init__(self, client_id: str, api_key: str, base_url: str, mock_mode: bool = True):
        self.client_id = client_id.strip()
        self.api_key = api_key.strip()
        self.base_url = base_url.rstrip("/")
        self.mock_mode = mock_mode
        self.logger = setup_logger()

    def _headers(self) -> Dict[str, str]:
        return {
            "Client-Id": self.client_id,
            "Api-Key": self.api_key,
            "Content-Type": "application/json",
        }

    def _log(self, endpoint: str, payload: Dict, result: Dict) -> None:
        safe_payload = dict(payload or {})
        self.logger.info(
            "endpoint=%s client_id=%s api_key=%s payload=%s result=%s",
            endpoint,
            self.client_id,
            mask_secret(self.api_key),
            json.dumps(safe_payload, ensure_ascii=False, default=str)[:5000],
            json.dumps(result, ensure_ascii=False, default=str)[:5000],
        )

    def _request(self, endpoint: str, payload: Dict) -> Dict:
        if self.mock_mode:
            result = {
                "success": True,
                "mock": True,
                "endpoint": endpoint,
                "task_id": f"mock-{datetime.now().strftime('%Y%m%d%H%M%S')}",
                "response": {"message": "mock 模式：未真实调用 Ozon API"},
            }
            self._log(endpoint, payload, result)
            return result

        if not self.client_id or not self.api_key:
            result = {"success": False, "error": "缺少 OZON_CLIENT_ID 或 OZON_API_KEY"}
            self._log(endpoint, payload, result)
            return result

        try:
            response = requests.post(
                f"{self.base_url}{endpoint}",
                headers=self._headers(),
                json=payload,
                timeout=60,
            )
            try:
                data = response.json()
            except ValueError:
                data = {"text": response.text}

            result = {
                "success": response.ok,
                "status_code": response.status_code,
                "endpoint": endpoint,
                "response": data,
                "task_id": data.get("result", {}).get("task_id") if isinstance(data, dict) else None,
            }
            if not response.ok:
                result["error"] = data
        except requests.RequestException as exc:
            result = {"success": False, "endpoint": endpoint, "error": str(exc)}

        self._log(endpoint, payload, result)
        return result

    def test_connection(self) -> Dict:
        # 使用商品列表做轻量权限测试；mock 模式不会真实请求。
        # 属性接口需要有效类目和 limit，拿来做通用连通测试容易误报。
        return self._request("/v3/product/list", {"filter": {"visibility": "ALL"}, "last_id": "", "limit": 1})

    def import_products(self, products: List[Dict]) -> Dict:
        return self._request("/v3/product/import", {"items": products})

    def import_pictures(self, pictures: List[Dict]) -> Dict:
        return self._request("/v1/product/pictures/import", {"pictures": pictures})

    def update_prices(self, prices: List[Dict]) -> Dict:
        return self._request("/v1/product/import/prices", {"prices": prices})

    def update_stocks(self, stocks: List[Dict]) -> Dict:
        # TODO: Ozon 店铺仓库库存可能需要 warehouse_id，请在 Ozon 后台/官方文档确认。
        return self._request("/v2/products/stocks", {"stocks": stocks})

    def get_import_info(self, task_id: str) -> Dict:
        return self._request("/v1/product/import/info", {"task_id": task_id})

    def get_category_attributes(self, description_category_id: int | str, type_id: int | str) -> Dict:
        return self._request(
            "/v4/product/info/attributes",
            {
                "filter": {
                    "description_category_id": int(description_category_id),
                    "type_id": int(type_id),
                },
                "limit": 1000,
            },
        )

    def get_description_category_tree(self, language: str = "ZH_HANS") -> Dict:
        return self._request("/v1/description-category/tree", {"language": language})


def products_to_import_payload(products: List[Dict]) -> List[Dict]:
    payload = []
    for product in products:
        try:
            attributes = json.loads(product.get("attributes_json") or "{}")
        except (TypeError, json.JSONDecodeError):
            attributes = {}
        item = {
            "offer_id": str(product.get("offer_id", "")),
            "name": str(product.get("name_ru", "")),
            "description": str(product.get("description_ru", "")),
            "barcode": str(product.get("barcode", "") or ""),
            # TODO: 按 Ozon 官方类目属性要求补充 complex_attributes / pdf_list / rich_content_json 等字段。
        }

        optional_int_fields = ["description_category_id", "type_id", "depth", "width", "height", "weight"]
        for field in optional_int_fields:
            value = product.get(field, "")
            if str(value).strip() != "":
                try:
                    item[field] = int(float(value))
                except (TypeError, ValueError):
                    pass

        optional_text_fields = ["price", "old_price"]
        for field in optional_text_fields:
            value = product.get(field, "")
            if str(value).strip() != "":
                item[field] = str(value)

        if str(product.get("currency_code", "")).strip():
            item["currency_code"] = product.get("currency_code")
        elif str(product.get("price", "")).strip():
            item["currency_code"] = "RUB"

        if str(product.get("vat", "")).strip() != "":
            item["vat"] = str(product.get("vat"))

        if any(str(product.get(field, "")).strip() for field in ["depth", "width", "height"]):
            item["dimension_unit"] = "mm"
        if str(product.get("weight", "")).strip():
            item["weight_unit"] = "g"

        parsed_attributes = attributes if isinstance(attributes, list) else attributes.get("attributes", [])
        if parsed_attributes:
            item["attributes"] = parsed_attributes

        image_urls = product.get("_api_image_urls", [])
        if image_urls:
            item["images"] = image_urls

        payload.append(item)
    return payload


def products_to_price_payload(products: List[Dict]) -> List[Dict]:
    return [
        {
            "offer_id": str(product.get("offer_id", "")),
            "price": str(product.get("price", "")),
            "old_price": str(product.get("old_price", "") or ""),
            "currency_code": product.get("currency_code") or "RUB",
            "vat": str(product.get("vat", 0) if product.get("vat", "") != "" else 0),
        }
        for product in products
    ]


def products_to_stock_payload(products: List[Dict]) -> List[Dict]:
    payload = []
    for product in products:
        item = {"offer_id": str(product.get("offer_id", ""))}
        stock = product.get("stock", "")
        try:
            item["stock"] = int(float(stock)) if str(stock).strip() else 0
        except (TypeError, ValueError):
            item["stock"] = 0
        # TODO: 真实库存更新通常需要 warehouse_id，请确认后在 Excel 增加字段。
        payload.append(item)
    return payload


def products_to_picture_payload(products: List[Dict]) -> List[Dict]:
    return [
        {
            "offer_id": str(product.get("offer_id", "")),
            "images": product.get("_api_image_urls", []),
        }
        for product in products
        if product.get("_api_image_urls")
    ]
