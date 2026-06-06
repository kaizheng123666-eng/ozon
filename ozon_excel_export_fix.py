# -*- coding: utf-8 -*-
"""
Ozon Excel 导出修复补丁。

用法：
    from ozon_excel_export_fix import export_ozon_xlsx_from_template

    export_ozon_xlsx_from_template(
        template_path="狗狗服装 雨衣模板.xlsx",
        output_path="宠物服装_生成结果.xlsx",
        products=products,
        description_category_id="17030187",
        product_weight_g=200,
    )

核心原则：必须基于 Ozon 官方模板复制填写，不能自己新建表头、删列、改 sheet。
"""
from __future__ import annotations

from copy import copy
from pathlib import Path
import re
from typing import Any

from openpyxl import load_work