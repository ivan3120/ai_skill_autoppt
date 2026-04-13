# -*- coding: utf-8 -*-
"""
配置模块

提供Skill所需的配置管理功能。
"""

import os
import yaml
from pathlib import Path
from typing import Optional


class Config:
    """配置类"""

    _instance = None
    _config = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        if self._config is None:
            self._load()

    def _load(self):
        """加载配置"""
        # skill目录路径
        skill_dir = Path(__file__).parent.parent

        # 使用默认配置
        self._config = {
            "llm": {
                "api_key": os.environ.get("LLM_API_KEY", ""),
                "base_url": os.environ.get("LLM_BASE_URL", "https://api.minimaxi.com/anthropic"),
                "model": os.environ.get("LLM_MODEL", "MiniMax-M2.5"),
            },
            "storage": {
                "base_dir": str(skill_dir / "storage"),
                "template_dir": str(skill_dir / "templates"),
                "ppt_dir": str(skill_dir / "storage" / "ppt"),
                "excel_dir": str(skill_dir / "storage" / "excel"),
            },
            "skill": {
                "default_level": 1,
                "enable_progressive": True,
                "templates": str(skill_dir / "templates"),
                "assets": str(skill_dir / "assets"),
                "references": str(skill_dir / "references"),
            }
        }

    @classmethod
    def get(cls, key: str, default=None):
        """获取配置项"""
        instance = cls()
        keys = key.split(".")
        value = instance._config
        for k in keys:
            if isinstance(value, dict):
                value = value.get(k)
            else:
                return default
        return value if value is not None else default

    @classmethod
    def is_llm_available(cls) -> bool:
        """检查LLM是否可用"""
        return bool(cls.get("llm.api_key"))

    @classmethod
    def get_llm_config(cls):
        """获取LLM配置"""
        return LLMConfig()

    @property
    def llm_config(self):
        """获取LLM配置"""
        return self._config.get("llm", {})


class LLMConfig:
    """LLM配置类"""

    def __init__(self):
        self.api_key = Config.get("llm.api_key", "")
        self.base_url = Config.get("llm.base_url", "https://api.minimaxi.com/anthropic")
        self.model = Config.get("llm.model", "MiniMax-M2.5")
        self.temperature = 0.7
        self.max_tokens = 4096

    @property
    def is_available(self) -> bool:
        """检查是否可用"""
        return bool(self.api_key)
