# -*- coding: utf-8 -*-
"""
AutoPPTSkill - 主入口类

提供统一的Skill调用接口。
"""

import os
from typing import Dict, Optional
from pathlib import Path


class AutoPPTSkill:
    """
    云核心网评估报告生成Skill

    提供智能化生成云核心网网络架构评估PPT报告的能力。
    """

    def __init__(self, config_path: Optional[str] = None):
        """
        初始化Skill

        Args:
            config_path: 配置文件路径 (默认: ./config.yaml)
        """
        self.config_path = config_path or self._get_default_config_path()
        self._load_config()

    def _get_default_config_path(self) -> str:
        """获取默认配置路径"""
        # 首先检查当前目录
        if os.path.exists("./config.yaml"):
            return "./config.yaml"
        # 然后检查包内
        package_dir = Path(__file__).parent.parent
        config_path = package_dir / "config.yaml"
        if config_path.exists():
            return str(config_path)
        return "./config.yaml"

    def _load_config(self):
        """加载配置"""
        import yaml

        if os.path.exists(self.config_path):
            with open(self.config_path, "r", encoding="utf-8") as f:
                self.config = yaml.safe_load(f)
        else:
            self.config = {}

    def generate(
        self,
        excel_path: str,
        template: str = "default",
        dimensions: list = None,
        domains: list = None,
        use_llm: bool = False,
        output_path: Optional[str] = None
    ) -> Dict:
        """
        生成PPT报告

        Args:
            excel_path: Excel评估数据文件路径
            template: PPT模板名称
            dimensions: 评估维度列表
            domains: 产品域列表
            use_llm: 是否使用LLM增强
            output_path: 输出文件路径 (可选)

        Returns:
            Dict: 包含status、data、message的结果字典
        """
        from scripts.workflows.agent_manager import generate_ppt_workflow

        user_selections = {
            "dimensions": dimensions or ["MT0", "MT1", "MT2", "MT3", "MT4", "MT5", "MT6"],
            "domains": domains or [],
            "template": template,
            "use_llm": use_llm
        }

        result = generate_ppt_workflow(excel_path, user_selections)

        # 如果指定了输出路径，移动文件
        if output_path and result.get("status") == "success":
            ppt_path = result.get("data", {}).get("ppt_path")
            if ppt_path and ppt_path != output_path:
                import shutil
                shutil.copy(ppt_path, output_path)
                result["data"]["ppt_path"] = output_path

        return result

    def parse(self, excel_path: str) -> Dict:
        """
        解析Excel数据（不生成PPT）

        Args:
            excel_path: Excel文件路径

        Returns:
            Dict: 本体模型数据
        """
        from scripts.services.excel_parser import ExcelParser

        parser = ExcelParser(excel_path)
        parser.load()
        ontology = parser.parse_all()

        return ontology.to_dict()

    def list_templates(self) -> list:
        """
        列出可用模板

        Returns:
            list: 模板文件列表
        """
        from pathlib import Path

        template_dir = self._get_resource_path("templates")
        if not template_dir.exists():
            return []

        templates = list(template_dir.glob("*.pptx"))
        return [t.stem for t in templates]

    def _get_resource_path(self, resource: str) -> Path:
        """获取资源目录路径"""
        package_dir = Path(__file__).parent.parent
        return package_dir / "resources" / resource


# 快捷函数
def generate_ppt(
    excel_path: str,
    template: str = "default",
    output_path: Optional[str] = None
) -> Dict:
    """
    快捷函数：生成PPT报告

    Args:
        excel_path: Excel文件路径
        template: 模板名称
        output_path: 输出路径

    Returns:
        Dict: 结果
    """
    skill = AutoPPTSkill()
    return skill.generate(excel_path, template=template, output_path=output_path)
