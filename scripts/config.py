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

    # LLM 提供商枚举
    PROVIDER_MINIMAX = "minimax"
    PROVIDER_ANTHROPIC = "anthropic"
    PROVIDER_OPENAI = "openai"

    def __init__(self):
        # 自动检测 LLM 提供商
        self._detect_provider()

        # MiniMax 配置
        self.minimax_api_key = os.environ.get("MINIMAX_API_KEY") or os.environ.get("LLM_API_KEY", "")
        self.minimax_base_url = os.environ.get("MINIMAX_BASE_URL") or os.environ.get("LLM_BASE_URL", "https://api.minimax.com/v1")
        self.minimax_model = os.environ.get("MINIMAX_MODEL") or os.environ.get("LLM_MODEL", "MiniMax-M2.5")

        # Anthropic 配置
        self.anthropic_api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        self.anthropic_base_url = os.environ.get("ANTHROPIC_BASE_URL", "https://api.anthropic.com")
        self.anthropic_model = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-20250514")

        # OpenAI 配置
        self.openai_api_key = os.environ.get("OPENAI_API_KEY", "")
        self.openai_base_url = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")
        self.openai_model = os.environ.get("OPENAI_MODEL", "gpt-4o")

        # 统一配置（基于检测到的提供商）
        self.api_key = self._get_provider_api_key()
        self.base_url = self._get_provider_base_url()
        self.model = self._get_provider_model()

        self.temperature = 0.7
        self.max_tokens = 4096

    def _detect_provider(self):
        """自动检测 LLM 提供商"""
        # 优先级: MiniMax > Anthropic > OpenAI
        if os.environ.get("MINIMAX_API_KEY") or os.environ.get("LLM_API_KEY"):
            self.provider = self.PROVIDER_MINIMAX
        elif os.environ.get("ANTHROPIC_API_KEY"):
            self.provider = self.PROVIDER_ANTHROPIC
        elif os.environ.get("OPENAI_API_KEY"):
            self.provider = self.PROVIDER_OPENAI
        else:
            self.provider = None

    def _get_provider_api_key(self) -> str:
        """获取当前提供商的 API Key"""
        if self.provider == self.PROVIDER_MINIMAX:
            return self.minimax_api_key
        elif self.provider == self.PROVIDER_ANTHROPIC:
            return self.anthropic_api_key
        elif self.provider == self.PROVIDER_OPENAI:
            return self.openai_api_key
        return ""

    def _get_provider_base_url(self) -> str:
        """获取当前提供商的 Base URL"""
        if self.provider == self.PROVIDER_MINIMAX:
            return self.minimax_base_url
        elif self.provider == self.PROVIDER_ANTHROPIC:
            return self.anthropic_base_url
        elif self.provider == self.PROVIDER_OPENAI:
            return self.openai_base_url
        return "https://api.minimax.com/v1"

    def _get_provider_model(self) -> str:
        """获取当前提供商的模型"""
        if self.provider == self.PROVIDER_MINIMAX:
            return self.minimax_model
        elif self.provider == self.PROVIDER_ANTHROPIC:
            return self.anthropic_model
        elif self.provider == self.PROVIDER_OPENAI:
            return self.openai_model
        return "MiniMax-M2.5"

    @property
    def is_available(self) -> bool:
        """检查是否可用"""
        return bool(self.api_key)

    @property
    def provider_name(self) -> str:
        """获取提供商名称"""
        return self.provider or "none"
