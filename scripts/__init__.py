# -*- coding: utf-8 -*-
"""
ai_skill_autoppt - 云核心网评估报告生成Skill

提供智能化生成云核心网网络架构评估PPT报告的能力。
"""

__version__ = "1.0.0"
__author__ = "AutoPPT Team"

from scripts.workflows.agent_manager import generate_ppt_workflow
from scripts.skill import AutoPPTSkill

__all__ = [
    "AutoPPTSkill",
    "generate_ppt_workflow",
    "__version__",
]
