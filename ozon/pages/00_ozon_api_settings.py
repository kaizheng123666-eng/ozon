from __future__ import annotations

import os
from pathlib import Path
from typing import Dict

import streamlit as st


def env_path() -> Path:
    return Path.cwd() / ".env"


def read_env(path: Path) -> Dict[str, str]:
    values: Dict[str, str] = {}
    if not path.exists():
        return values
    try:
        for raw_line in path.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            values[key.strip()] = value.strip().strip('"').strip("'")
    except Exception:
        return values
    return values


def write_env(path: Path, updates: Dict[str, str]) -> None:
    current = read_env(path)
    current.update({key: str(value).strip() for key, value in updates.items()})
    ordered_keys = ["OZON_CLIENT_ID", "OZON_API_KEY", "OZON_BASE_URL"]
    other_keys = [key for key in current.keys() if key not in ordered_keys]
    lines = []
    for key in ordered_keys + other_keys:
        if key in current:
            lines.append(f"{key}={current[key]}")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    for key, value in updates.items():
        os.environ[key] = str(value).strip()


def mask_secret(value: str) -> str:
    value = str(value or "")
    if not value:
        return "未填写"
    if len(value) <= 6:
        return "*" * len(value)
    return value[:3] + "***" + value[-3:]


st.set_page_config(page_title="Ozon API 设置", layout="wide")
st.title("Ozon API 设置")
st.info("这里填写的是 Ozon 卖家后台的 Client-Id 和 Api-Key，不是 OpenAI API。信息只保存到本地 .env 文件，不会提交到 GitHub。")

path = env_path()
values = read_env(path)

client_id_default = os.getenv("OZON_CLIENT_ID", values.get("OZON_CLIENT_ID", ""))
api_key_default = os.getenv("OZON_API_KEY", values.get("OZON_API_KEY", ""))
base_url_default = os.getenv("OZON_BASE_URL", values.get("OZON_BASE_URL", "https://api-seller.ozon.ru"))

with st.form("ozon_api_settings_form"):
    client_id = st.text_input("Client-Id", value=client_id_default, help="Ozon Seller 后台 API 页面里的 Client-Id")
    api_key = st.text_input("Api-Key", value=api_key_default, type="password", help="Ozon Seller 后台创建的 API Key")
    base_url = st.text_input("Ozon API 地址", value=base_url_default or "https://api-seller.ozon.ru")
    submitted = st.form_submit_button("保存 API 设置", type="primary")

if submitted:
    if not client_id.strip() or not api_key.strip():
        st.error("Client-Id 和 Api-Key 都要填写。")
    else:
        write_env(
            path,
            {
                "OZON_CLIENT_ID": client_id,
                "OZON_API_KEY": api_key,
                "OZON_BASE_URL": base_url or "https://api-seller.ozon.ru",
            },
        )
        st.success(f"已保存到本地文件：{path}")
        st.caption("保存后请回到 Ozon API 上传页面，或刷新页面/重新运行 streamlit。")

st.subheader("当前读取状态")
col1, col2, col3 = st.columns(3)
col1.metric("Client-Id", mask_secret(client_id_default))
col2.metric("Api-Key", mask_secret(api_key_default))
col3.metric("API 地址", base_url_default or "默认")

st.warning("不要把 .env 文件上传到 GitHub。只提交 .env.example 示例文件。")
