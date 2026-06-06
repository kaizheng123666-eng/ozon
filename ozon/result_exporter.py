from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Dict, List

import pandas as pd

from config import OUTPUT_DIR


RESULT_COLUMNS = [
    "offer_id",
    "name_ru",
    "status",
    "step",
    "success",
    "task_id",
    "ozon_product_id",
    "sku",
    "error_message",
    "warning_message",
    "uploaded_at",
]


def export_results(results: List[Dict], output_path: str | Path | None = None) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if output_path is None:
        output_path = OUTPUT_DIR / f"upload_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    output = Path(output_path)

    rows = []
    for result in results:
        rows.append({col: result.get(col, "") for col in RESULT_COLUMNS})
    pd.DataFrame(rows, columns=RESULT_COLUMNS).to_excel(output, index=False)
    return output


def make_result(product: Dict, step: str, api_result: Dict | None = None, validation: Dict | None = None) -> Dict:
    api_result = api_result or {}
    validation = validation or {}
    response = api_result.get("response", {}) if isinstance(api_result, dict) else {}
    return {
        "offer_id": product.get("offer_id", ""),
        "name_ru": product.get("name_ru", ""),
        "status": validation.get("status", "已处理" if api_result.get("success") else "失败"),
        "step": step,
        "success": bool(api_result.get("success", validation.get("success", False))),
        "task_id": api_result.get("task_id", ""),
        "ozon_product_id": response.get("result", {}).get("product_id", "") if isinstance(response, dict) else "",
        "sku": response.get("result", {}).get("sku", "") if isinstance(response, dict) else "",
        "error_message": validation.get("error_message", "") or str(api_result.get("error", "")),
        "warning_message": validation.get("warning_message", ""),
        "uploaded_at": datetime.now().isoformat(timespec="seconds"),
    }

