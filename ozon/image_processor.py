from __future__ import annotations

import re
import shutil
from pathlib import Path
from typing import Dict, Iterable, List

from PIL import Image


SUPPORTED_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}


def list_product_images(folder: str | Path) -> List[Path]:
    root = Path(folder)
    if not root.exists():
        return []
    return sorted(
        [p for p in root.rglob("*") if p.is_file() and p.suffix.lower() in SUPPORTED_IMAGE_EXTENSIONS],
        key=lambda p: p.name.lower(),
    )


def safe_offer_id(value: str) -> str:
    text = re.sub(r"[^\w\-.]+", "-", str(value or "").strip(), flags=re.UNICODE)
    return text.strip("-") or "OZON-ITEM"


def _find_by_name(images: Iterable[Path], name: str) -> Path | None:
    if not name:
        return None
    lowered = name.lower()
    for path in images:
        if path.name.lower() == lowered:
            return path
    return None


def _save_as_png(source: Path, target: Path, max_side: int = 3000) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    with Image.open(source) as image:
        image = image.convert("RGB")
        image.thumbnail((max_side, max_side), Image.Resampling.LANCZOS)
        image.save(target, format="PNG", optimize=True)


def process_product_images(
    image_paths: List[str | Path],
    output_root: str | Path,
    offer_id: str,
    image_roles: Dict | None = None,
) -> Dict[str, List[str] | str]:
    images = [Path(path) for path in image_paths]
    roles = image_roles or {}
    clean_offer_id = safe_offer_id(offer_id)
    target_dir = Path(output_root) / clean_offer_id
    target_dir.mkdir(parents=True, exist_ok=True)

    primary_source = _find_by_name(images, roles.get("primary_image", "")) or (images[0] if images else None)
    size_source = _find_by_name(images, roles.get("size_chart", ""))

    used: set[Path] = set()
    primary_name = ""
    if primary_source:
        primary_target = target_dir / f"{clean_offer_id}_main.png"
        _save_as_png(primary_source, primary_target)
        primary_name = primary_target.name
        used.add(primary_source)

    extra_names: List[str] = []
    size_name = ""
    if size_source and size_source not in used:
        size_target = target_dir / f"{clean_offer_id}_size.png"
        _save_as_png(size_source, size_target)
        size_name = size_target.name
        extra_names.append(size_target.name)
        used.add(size_source)

    role_names = list(roles.get("additional_images", []) or []) + list(roles.get("detail_images", []) or [])
    ordered_sources: List[Path] = []
    for name in role_names:
        path = _find_by_name(images, name)
        if path and path not in ordered_sources and path not in used:
            ordered_sources.append(path)
    for path in images:
        if path not in ordered_sources and path not in used:
            ordered_sources.append(path)

    for index, source in enumerate(ordered_sources, start=1):
        target = target_dir / f"{clean_offer_id}_{index}.png"
        _save_as_png(source, target)
        extra_names.append(target.name)
        used.add(source)

    original_dir = target_dir / "original"
    original_dir.mkdir(exist_ok=True)
    for source in images:
        shutil.copy2(source, original_dir / source.name)

    return {
        "folder": str(target_dir),
        "primary_image": primary_name,
        "additional_images": extra_names,
        "size_image": size_name,
    }
