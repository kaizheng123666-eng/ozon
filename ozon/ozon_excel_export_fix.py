# -*- coding: utf-8 -*-
"""
Ozon Excel 导出修复补丁。

核心规则：
1. 必须基于 Ozon 官方模板填写，不能重新生成表头。
2. 只写“模板”sheet 第 5 行以后，不改 configs / info / validation 的结构。
3. 材料/材质列留空。
4. “一个商品中的件数”和“原厂包装数量”固定写 1。
5. “商品重量，克”同步手动填写的重量。
6. “颜色名称”自动把中文颜色转成俄语。
7. 支持每个颜色单独填写主图链接和附图链接。
8. description_category_id 会同步到 configs 的 DESCRIPTION_CATEGORY_ID。
"""
from __future__ import annotations

from copy import copy
from pathlib import Path
import re
from typing import Any

from openpyxl import load_workbook


COLOR_CN_TO_RU: dict