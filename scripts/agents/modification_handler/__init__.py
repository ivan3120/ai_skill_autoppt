# -*- coding: utf-8 -*-
"""
ModificationHandler Agent - 修改处理器

独立解耦的Agent模块，负责解析用户的自然语言修改指令。
"""

import re
from typing import Dict, List
from scripts.agents.base import BaseAgent, AgentResult


class ModificationHandlerAgent(BaseAgent):
    """
    修改处理器Agent
    任务：解析用户的自然语言修改指令，如"把标题改为..."、"在第3页添加图表..."
    """

    def __init__(self):
        super().__init__("ModificationHandler", "修改处理器", "modification_handler")

    def execute(self, input_data: Dict) -> AgentResult:
        """
        处理修改指令
        """
        modification_text = input_data.get("modification_text", "")
        current_content = input_data.get("current_content", [])

        if self._has_llm():
            modification_result = self._llm_modify(modification_text, current_content)
        else:
            modification_result = self._rule_based_modification(modification_text, current_content)

        if modification_result.get("status") == "success":
            modified_content = self._apply_modifications(current_content, modification_result)
            message = f"修改完成，共处理 {len(modification_result.get('modifications', []))} 项"
        else:
            modified_content = current_content
            message = modification_result.get("message", "修改失败")

        return AgentResult(
            status=modification_result.get("status", "partial"),
            data={"modified_content": modified_content, "modification_result": modification_result},
            message=message
        )

    def _llm_modify(self, modification_text: str, current_content: List[Dict]) -> Dict:
        """使用LLM解析和执行修改"""
        content_summary = self._build_content_summary(current_content)

        user_message = self._load_llm_prompt(
            self.llm_prompt_file,
            modification_text=modification_text,
            current_content=content_summary
        )

        try:
            llm_result = self.invoke_llm(user_message)
            self._safe_print(f"[ModificationHandler] LLM: {llm_result[:200]}...")
            modification_result = self._parse_modification(llm_result)
        except Exception as e:
            self._safe_print(f"[ModificationHandler] LLM modification failed: {e}")
            modification_result = self._rule_based_modification(modification_text, current_content)

        return modification_result

    def _build_content_summary(self, content: List[Dict]) -> str:
        """构建内容摘要用于修改"""
        summary_lines = []
        for i, slide in enumerate(content):
            slide_type = slide.get("type", "")
            title = slide.get("title", "")
            summary_lines.append(f"第{i+1}页 [{slide_type}] {title}")

        return "\n".join(summary_lines)

    def _rule_based_modification(self, text: str, content: List[Dict]) -> Dict:
        """基于规则的修改解析"""
        modifications = []

        page_match = re.search(r"第(\d+)页", text)
        page = int(page_match.group(1)) - 1 if page_match else 0

        if page < 0 or page >= len(content):
            return {"status": "failed", "message": f"页码 {page + 1} 不存在", "modifications": []}

        if "改标题" in text or "标题改为" in text:
            new_title_match = re.search(r"(?:标题改为|改成)(.+?)(?:$|,|。)", text)
            if new_title_match:
                modifications.append({
                    "page": page, "operation": "modify_title", "target": "标题",
                    "old_value": content[page].get("title", ""),
                    "new_value": new_title_match.group(1).strip()
                })

        elif "添加" in text or "增加" in text:
            new_content_match = re.search(r"(?:添加|增加)(.+?)(?:$|,|。)", text)
            if new_content_match:
                modifications.append({
                    "page": page, "operation": "add_content", "target": "内容",
                    "new_value": new_content_match.group(1).strip()
                })

        return {
            "status": "success" if modifications else "partial",
            "message": "修改完成" if modifications else "未能解析修改指令",
            "modifications": modifications
        }

    def _parse_modification(self, llm_result: str) -> Dict:
        """解析LLM的修改结果"""
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
            self._safe_print(f"[ModificationHandler] Parse failed: {e}")

        return {"status": "failed", "message": "无法解析修改指令", "modifications": []}

    def _apply_modifications(self, content: List[Dict], modification_result: Dict) -> List[Dict]:
        """应用修改到PPT内容"""
        modified = [slide.copy() for slide in content]
        modifications = modification_result.get("modifications", [])

        for mod in modifications:
            page = mod.get("page", 0)
            operation = mod.get("operation", "")
            new_value = mod.get("new_value", "")

            if page < 0 or page >= len(modified):
                continue

            if operation == "modify_title":
                modified[page]["title"] = new_value
            elif operation == "modify_content":
                if "content" not in modified[page]:
                    modified[page]["content"] = {}
                modified[page]["content"]["文本"] = new_value
            elif operation == "add_content":
                if "content" not in modified[page]:
                    modified[page]["content"] = {}
                existing = modified[page]["content"].get("文本", "")
                modified[page]["content"]["文本"] = existing + "\n" + new_value

        return modified
