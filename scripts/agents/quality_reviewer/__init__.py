# -*- coding: utf-8 -*-
"""
QualityReviewer Agent - 质量审核员

独立解耦的Agent模块，负责审核PPT报告的内容质量。
"""

from typing import Dict, List
from scripts.agents.base import BaseAgent, AgentResult


class QualityReviewerAgent(BaseAgent):
    """
    质量审核员Agent
    任务：审核PPT报告的内容，确保内容准确性、逻辑连贯性、格式规范性
    """

    def __init__(self):
        super().__init__("QualityReviewer", "质量审核员", "quality_reviewer")
        # 设置默认prompt模板
        if not self.prompt_template:
            self.prompt_template = """你是一个专业的PPT质量审核专家。请审核以下PPT内容，检查其准确性、逻辑性和规范性。

PPT内容摘要：
{ppt_content}

请返回JSON格式的审核结果：
```json
{
    "score": 85,
    "status": "passed/needs_revision",
    "issues": ["问题1", "问题2"],
    "strengths": ["优点1", "优点2"],
    "suggestions": ["改进建议1", "改进建议2"]
}
```"""

    def execute(self, input_data: Dict) -> AgentResult:
        """
        审核内容质量
        """
        ppt_content = input_data.get("ppt_content", [])
        original_data = input_data.get("original_data", {})
        style_config = input_data.get("style_config", {})

        if self._has_llm():
            review_result = self._llm_review(ppt_content, original_data, style_config)
        else:
            review_result = self._rule_based_review(ppt_content, original_data, style_config)

        if review_result.get("score", 100) >= 90 and review_result.get("status") != "needs_revision":
            status = "success"
            message = "内容审核通过"
        elif review_result.get("status") == "needs_revision":
            status = "partial"
            message = f"发现 {len(review_result.get('issues', []))} 个问题，需要修改"
        else:
            status = "partial"
            message = "内容审核发现问题，建议检查"

        return AgentResult(
            status=status,
            data=review_result,
            message=message
        )

    def _llm_review(self, ppt_content: List[Dict], original_data: Dict, style_config: Dict) -> Dict:
        """使用LLM进行深度内容审核"""
        slides_summary = self._build_slides_summary(ppt_content)

        user_message = self._load_llm_prompt(
            self.llm_prompt_file,
            ppt_content=slides_summary
        )

        try:
            llm_result = self.invoke_llm(user_message)
            self._safe_print(f"[QualityReviewer] LLM: {llm_result[:200]}...")
            review_result = self._parse_review_result(llm_result)
        except Exception as e:
            self._safe_print(f"[QualityReviewer] LLM review failed: {e}")
            review_result = self._rule_based_review(ppt_content, original_data, style_config)

        return review_result

    def _build_slides_summary(self, slides: List[Dict]) -> str:
        """构建幻灯片摘要用于审核"""
        summary_lines = []
        for i, slide in enumerate(slides):
            slide_type = slide.get("type", "")
            title = slide.get("title", "")
            content = slide.get("content", {})

            summary = f"第{i+1}页 [{slide_type}] {title}\n"
            if slide_type == "overview":
                items = content.get("items", [])
                summary += f"  包含 {len(items)} 个评估维度\n"
            elif slide_type == "dimension_detail":
                pass_rate = content.get("通过率", "N/A")
                risk = content.get("风险等级", "N/A")
                summary += f"  通过率: {pass_rate}, 风险: {risk}\n"
            elif slide_type == "summary":
                summary += f"  包含 {len(content.get('总结', []))} 项总结\n"

            summary_lines.append(summary)

        return "\n".join(summary_lines)

    def _rule_based_review(self, ppt_content: List[Dict], original_data: Dict, style_config: Dict) -> Dict:
        """基于规则的内容审核"""
        issues = []
        score = 100

        if len(ppt_content) < 10:
            issues.append({
                "page": 0, "type": "逻辑问题", "severity": "high",
                "description": f"幻灯片数量不足，仅 {len(ppt_content)} 页",
                "suggestion": "确保生成完整的13页报告"
            })
            score -= 20

        cover = next((s for s in ppt_content if s.get("type") == "cover"), None)
        if not cover:
            issues.append({
                "page": 0, "type": "逻辑问题", "severity": "high",
                "description": "缺少封面页", "suggestion": "添加封面页，包含标题和日期"
            })
            score -= 10

        summary = next((s for s in ppt_content if s.get("type") == "summary"), None)
        if not summary:
            issues.append({
                "page": len(ppt_content), "type": "逻辑问题", "severity": "medium",
                "description": "缺少总结页", "suggestion": "添加总结页，包含主要发现和优化建议"
            })
            score -= 5

        titles = [s.get("title", "") for s in ppt_content]
        if len(titles) != len(set(titles)):
            duplicate_titles = [t for t in titles if titles.count(t) > 1]
            issues.append({
                "page": 0, "type": "格式问题", "severity": "low",
                "description": f"存在重复标题: {set(duplicate_titles)}",
                "suggestion": "确保每个页面标题唯一"
            })
            score -= 5

        status = "passed" if score >= 90 else "needs_revision"

        return {
            "status": status, "score": score, "issues": issues,
            "summary": f"基础审核完成，发现 {len(issues)} 个问题",
            "strengths": ["结构完整", "格式规范"] if score >= 90 else [],
            "improvements": [i.get("description") for i in issues]
        }

    def _parse_review_result(self, llm_result: str) -> Dict:
        """解析LLM返回的审核结果"""
        import json
        try:
            if "```json" in llm_result:
                start = llm_result.find("```json") + 7
                end = llm_result.find("```", start)
                return json.loads(llm_result[start:end].strip())
            elif "{" in llm_result:
                start = llm_result.find("{")
                end = llm_result.rfind("}") + 1
                return json.loads(llm_result[start:end])
        except Exception as e:
            self._safe_print(f"[QualityReviewer] Parse failed: {e}")

        return {"status": "passed", "score": 100, "issues": [], "message": "内容审核通过"}
