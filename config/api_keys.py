"""API key management for Mistral — pooled with rotation."""

import os
from typing import Optional
from utils.logging_setup import get_logger

logger = get_logger('config.api_keys')

CONFIG_FILE_PATH = 'config.ini'

HARDCODED_KEYS = []


def _collect_keys() -> list[str]:
    """Collect all available API keys from all sources."""
    keys = []

    # 1. Streamlit secrets
    try:
        import streamlit as st
        if hasattr(st, 'secrets'):
            for key_name in ['MISTRAL_API_KEY', 'MISTRAL_API_KEY_2', 'MISTRAL_API_KEY_3']:
                if key_name in st.secrets:
                    val = st.secrets[key_name]
                    if val and val not in keys:
                        keys.append(val)
    except Exception:
        pass

    # 2. Environment variables
    for suffix in ['', '_2', '_3', '_4', '_5', '_6', '_7', '_8', '_9']:
        env_key = os.environ.get(f'MISTRAL_API_KEY{suffix}')
        if env_key and env_key not in keys:
            keys.append(env_key)

    # 3. config.ini
    if os.path.exists(CONFIG_FILE_PATH):
        try:
            import configparser
            config = configparser.ConfigParser()
            config.read(CONFIG_FILE_PATH)
            key = config['API']['MISTRAL_API_KEY']
            if key and key != 'YOUR_API_KEY_HERE' and key not in keys:
                keys.append(key)
        except Exception:
            pass

    # 4. Hardcoded fallbacks
    for key in HARDCODED_KEYS:
        if key not in keys:
            keys.append(key)

    return keys


class ApiKeyPool:
    """Pool of API keys with round-robin rotation."""

    def __init__(self):
        self._index = 0
        self._keys = _collect_keys()

    def refresh(self):
        self._keys = _collect_keys()

    def get_current(self) -> Optional[str]:
        if not self._keys:
            return None
        return self._keys[self._index % len(self._keys)]

    def rotate(self) -> Optional[str]:
        if not self._keys:
            return None
        self._index = (self._index + 1) % len(self._keys)
        return self._keys[self._index]

    @property
    def size(self) -> int:
        return len(self._keys)


pool = ApiKeyPool()
