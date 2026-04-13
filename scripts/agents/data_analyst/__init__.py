# -*- coding: utf-8 -*-
"""
DataAnalyst Agent - 数据分析师

独立解耦的Agent模块，负责解析Excel文件并构建本体模型。
"""

import json
from typing import Dict
from scripts.agents.base import BaseAgent, AgentResult


class DataAnalystAgent(BaseAgent):
    """
    数据分析师Agent

    角色：资深数据分析师，专门负责分析云核心网网络架构评估数据
    任务：解析Excel文件，提取产品域、网元、容灾组等数据，构建本体模型
    """

    def __init__(self):
        super().__init__("DataAnalyst", "数据分析师", "data_analyst")

    def execute(self, input_data: Dict) -> AgentResult:
        """
        解析Excel数据

        输入示例:
        {
            "excel_path": "/path/to/excel.xlsx",
            "user_selections": ["MT0", "MT1"]
        }

        输出示例:
        {
            "status": "success",
            "data": {...},
            "message": f"Parsed {ontology_model.domain_count} product domains"
        }
        """
        excel_path = input_data.get("excel_path")
        if not excel_path:
            return AgentResult(status="error", message="未提供Excel文件路径")

        # 调用Excel解析服务
        from scripts.services.excel_parser import ExcelParser

        try:
            parser = ExcelParser(excel_path)
            parser.load()
            ontology_model = parser.parse_all()

            # 使用LLM增强解析（可选）
            if self._has_llm() and input_data.get("use_llm", False):
                # 从prompt文件加载LLM指令
                user_message = self._load_llm_prompt(
                    self.llm_prompt_file,
                    ontology_data=json.dumps(ontology_model.to_dict(), ensure_ascii=False, indent=2)
                )
                llm_result = self.invoke_llm(user_message, context=ontology_model.to_dict())
                message = f"Parsed {ontology_model.domain_count} domains, LLM analysis: {llm_result[:100]}..."
            else:
                message = f"Parsed {ontology_model.domain_count} product domains"

            return AgentResult(
                status="success",
                data=ontology_model.to_dict(),
                message=message
            )
        except Exception as e:
            return AgentResult(status="error", message=f"解析失败: {str(e)}")
