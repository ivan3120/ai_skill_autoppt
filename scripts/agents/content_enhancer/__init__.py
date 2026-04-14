# -*- coding: utf-8 -*-
"""
ContentEnhancer Agent - 内容增强器

专门负责使用 LLM 生成专业的分析结论、优化建议等内容。
作为 ContentPlanner 的补充，提供智能化内容生成。
"""

import json
from typing import Dict, List, Optional
from scripts.agents.base import BaseAgent, AgentResult


class ContentEnhancerAgent(BaseAgent):
    """
    内容增强器 Agent

    职责：
    - 生成专业的分析结论
    - 识别关键风险点
    - 提供可执行的优化建议
    - 生成整体总结
    """

    def __init__(self):
        super().__init__("ContentEnhancer", "内容增强器", "content_enhancer")

    def execute(self, input_data: Dict) -> AgentResult:
        """
        生成增强内容

        输入：
        {
            "parsed_data": {...},      # 解析后的数据
            "dimensions": ["MT0", ...], # 选中的维度
            "dimension_summary": {...} # 维度汇总数据
        }

        输出：
        {
            "dimension_analysis": [...],  # 每个维度的分析
            "overall_summary": "...",     # 整体总结
            "priorities": [...],          # 优先级建议
            "next_steps": "..."           # 下一步计划
        }
        """
        parsed_data = input_data.get("parsed_data", {})
        dimensions = input_data.get("dimensions", [])
        dimension_summary = input_data.get("dimension_summary", {})

        if not self._has_llm():
            return AgentResult(
                status="error",
                message="LLM 未配置，无法生成内容"
            )

        # 构建 prompt
        user_message = self._build_enhancement_prompt(
            parsed_data, dimensions, dimension_summary
        )

        # 调用 LLM
        llm_result = self.invoke_llm(user_message, context=parsed_data)
        self._safe_print(f"[ContentEnhancer] LLM: {llm_result[:150]}...")

        # 解析结果
        enhanced_content = self._parse_enhancement_result(llm_result)

        return AgentResult(
            status="success",
            data=enhanced_content,
            message="内容增强完成"
        )

    def _build_enhancement_prompt(
        self,
        data: Dict,
        dimensions: List[str],
        dimension_summary: Dict
    ) -> str:
        """构建内容增强的 prompt"""

        import json

        overview = data.get("整体评估概览", [])
        product_domains = data.get("产品域列表", [])

        # 计算整体通过率
        total_pass = sum(s.get("通过数", 0) for s in dimension_summary.values())
        total_all = total_pass + sum(
            s.get("待改进数", 0) + s.get("不通过数", 0)
            for s in dimension_summary.values()
        )
        overall_rate = f"{int(total_pass * 100 / total_all)}%" if total_all > 0 else "0%"

        # 识别高风险维度
        high_risk = [
            dim for dim, s in dimension_summary.items()
            if s.get("风险") == "高"
        ]

        prompt = f"""你是一位资深的云核心网架构评估专家。请根据以下评估数据，生成专业的分析内容。

## 评估维度
{', '.join(dimensions)}

## 产品域
{json.dumps(product_domains[:5], ensure_ascii=False, indent=2)[:800]}

## 整体评估概览
{json.dumps(overview[:15], ensure_ascii=False, indent=2)[:1000]}

## 整体通过率
{overall_rate}

## 高风险维度
{', '.join(high_risk) if high_risk else '无'}

## 要求
请生成以下内容（JSON格式）：

1. 每个维度的分析（2-3句话的专业结论）
2. 关键发现（每个维度 1-3 条）
3. 风险描述（用专业语言描述）
4. 优化建议（可执行的建议）
5. 整体总结（3-4句话）
6. 优先级排序的优化项

输出格式：
```json
{{
  "dimension_analysis": [
    {{
      "dimension": "MT0",
      "conclusion": "分析结论",
      "key_findings": ["发现1", "发现2"],
      "risks": ["风险描述"],
      "suggestions": ["优化建议"]
    }}
  ],
  "overall_summary": "整体总结",
  "priorities": [
    {{"priority": "高", "issue": "问题", "suggestion": "建议"}}
  ],
  "next_steps": "下一步计划"
}}
```"""

        return prompt

    def _parse_enhancement_result(self, llm_result: str) -> Dict:
        """解析 LLM 返回的增强内容"""

        # 尝试解析 JSON
        try:
            # 从 markdown 代码块中提取
            if "```json" in llm_result:
                start = llm_result.find("```json") + 7
                end = llm_result.find("```", start)
                if end > start:
                    json_str = llm_result[start:end].strip()
                    return json.loads(json_str)
            elif "{" in llm_result:
                start = llm_result.find("{")
                end = llm_result.rfind("}") + 1
                if end > start:
                    json_str = llm_result[start:end]
                    return json.loads(json_str)
        except json.JSONDecodeError as e:
            self._safe_print(f"[ContentEnhancer] JSON parse error: {e}")

        # 解析失败，返回默认结构
        return {
            "dimension_analysis": [],
            "overall_summary": "评估数据处理中...",
            "priorities": [],
            "next_steps": "继续分析中..."
        }

    def generate_summary(self, data: Dict, dimension_summary: Dict) -> AgentResult:
        """
        专门生成总结页内容
        """
        # 计算整体通过率
        total_pass = sum(s.get("通过数", 0) for s in dimension_summary.values())
        total_all = total_pass + sum(
            s.get("待改进数", 0) + s.get("不通过数", 0)
            for s in dimension_summary.values()
        )
        overall_rate = f"{int(total_pass * 100 / total_all)}%" if total_all > 0 else "0%"

        # 识别高风险维度
        high_risk_dims = [
            dim for dim, s in dimension_summary.items()
            if s.get("风险") == "高"
        ]

        if not self._has_llm():
            # 规则引擎 fallback
            summary = f"整体评估结果: {overall_rate}通过率"
            if high_risk_dims:
                summary += f"，{', '.join(high_risk_dims)}存在较高风险，需重点关注"

            return AgentResult(
                status="success",
                data={
                    "overall": summary,
                    "priorities": [
                        {"优先级": "高", "问题": dim, "建议": "需优化"}
                        for dim in high_risk_dims[:3]
                    ] if high_risk_dims else [],
                    "next_steps": "建议分阶段实施优化，优先处理高风险项"
                },
                message="使用规则引擎生成总结"
            )

        # 使用 LLM 生成
        prompt = f"""基于以下评估数据，生成专业的总结与优化建议：

整体通过率: {overall_rate}
高风险维度: {', '.join(high_risk_dims) if high_risk_dims else '无'}
维度详情: {dimension_summary}

请生成：
1. 整体总结（3-4句话）
2. 优先级排序的优化建议
3. 下一步行动计划

输出 JSON 格式：
```json
{{
  "overall": "总结内容",
  "priorities": [{{"优先级": "高", "问题": "问题", "建议": "建议"}}],
  "next_steps": "下一步计划"
}}
```"""

        llm_result = self.invoke_llm(prompt, context=data)
        return self._parse_enhancement_result(llm_result)
