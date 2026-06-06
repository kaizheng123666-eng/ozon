from __future__ import annotations

from pathlib import Path
from typing import Dict, Iterable, List
from urllib.parse import urlparse


SUPPORTED_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}


def is_url(value: str) -> bool:
    try:
        parsed = urlparse(value)
        return parsed.scheme in {"http", "https"} and bool(parsed.netloc)
    except Exception:
        return False


def split_images(value) -> List[str]:
    if not value:
        return []
    if isinstance(value, str):
        return [item.strip() for item in value.replace("\n", ",").split(",") if item.strip()]
    return [str(value).strip()]


def _iter_images(root: Path) -> Iterable[Path]:
    if not root.exists():
        return []
    return (
        p
        for p in root.rglob("*")
        if p.is_file() and p.suffix.lower() in SUPPORTED_IMAGE_EXTENSIONS
    )


def build_image_index(image_root: str | Path) -> Dict[str, List[Path]]:
    root = Path(image_root)
    index: Dict[str, List[Path]] = {}
    for path in _iter_images(root):
        index.setdefault(path.name.lower(), []).append(path)
    return index


def resolve_image(reference: str, image_root: str | Path, image_index: Dict[str, List[Path]]) -> Dict:
    ref = str(reference or "").strip()
    if not ref:
        return {
            "reference": ref,
            "exists": False,
            "is_url": False,
            "resolved": "",
            "api_url": "",
            "error": "图片字段为空",
        }

    if is_url(ref):
        return {
            "reference": ref,
            "exists": True,
            "is_url": True,
            "resolved": ref,
            "api_url": ref,
            "error": "",
        }

    candidate = Path(ref)
    root = Path(image_root)
    candidates = []
    if candidate.is_absolute():
        candidates.append(candidate)
    else:
        candidates.append(root / candidate)
        candidates.extend(image_index.get(candidate.name.lower(), []))

    for path in candidates:
        if path.exists() and path.is_file() and path.suffix.lower() in SUPPORTED_IMAGE_EXTENSIONS:
            return {
                "reference": ref,
                "exists": True,
                "is_url": False,
                "resolved": str(path),
                "api_url": "",
                "error": "",
            }

    return {
        "reference": ref,
        "exists": False,
        "is_url": False,
        "resolved": "",
        "api_url": "",
        "error": f"找不到图片：{ref}",
    }


def map_product_images(products: List[Dict], image_root: str | Path) -> List[Dict]:
    image_index = build_image_index(image_root)
    mapped = []

    for product in products:
        row = dict(product)
        primary = resolve_image(row.get("primary_image", ""), image_root, image_index)
        extras = [
            resolve_image(item, image_root, image_index)
            for item in split_images(row.get("images", ""))
        ]

        all_images = [primary] + extras
        row["_primary_image_resolved"] = primary
        row["_images_resolved"] = extras
        row["_all_images_resolved"] = all_images
        row["_api_image_urls"] = [item["api_url"] for item in all_images if item.get("api_url")]
        row["_local_image_paths"] = [item["resolved"] for item in all_images if item.get("exists") and not item.get("is_url")]
        mapped.append(row)

    return mapped

