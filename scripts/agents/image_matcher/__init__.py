# -*- coding: utf-8 -*-
"""
ImageMatcher Agent - 图片匹配师

独立解耦的Agent模块，负责从图片知识库中匹配原理性介绍图片。
"""

import os
from typing import Dict, List, Optional
from scripts.agents.base import BaseAgent, AgentResult


class ImageMatcherAgent(BaseAgent):
    """
    图片匹配师Agent
    任务：根据评估项从图片知识库中匹配合适的原理性介绍图片
    """

    def __init__(self):
        super().__init__("ImageMatcher", "图片匹配师", "image_matcher")
        # 设置默认prompt模板（当没有prompt文件时使用）
        if not self.prompt_template:
            self.prompt_template = """你是一个专业的图片匹配助手。根据以下信息，从图片库中匹配最合适的图片。

请根据以下信息匹配图片：
- 评估维度：{dimension}
- 评估项：{item}
- 产品域：{product_domain}
- 图片库路径：{image_base_path}

请返回JSON格式的结果：
```json
{
    "image_path": "图片路径或关键词",
    "reason": "匹配原因"
}
```"""

    def execute(self, input_data: Dict) -> AgentResult:
        """
        匹配原理性图片
        """
        dimension = input_data.get("dimension", "")
        item = input_data.get("item", "")
        product_domain = input_data.get("product_domain", "")
        slides = input_data.get("slides", [])

        image_mapping = {}

        if self._has_llm():
            image_mapping = self._llm_match_images(dimension, item, product_domain, slides)
        else:
            image_mapping = self._rule_based_match_all(slides)

        return AgentResult(
            status="success",
            data={"image_mapping": image_mapping},
            message=f"图片匹配完成，共匹配 {len(image_mapping)} 张图片"
        )

    def _llm_match_images(self, dimension: str, item: str, product_domain: str, slides: List[Dict]) -> Dict:
        """使用LLM为每个需要图片的幻灯片匹配图片"""
        image_mapping = {}

        for i, slide in enumerate(slides):
            slide_type = slide.get("type", "")
            slide_title = slide.get("title", "")

            needs_image = slide_type in ["dimension_detail", "background", "intro"]
            if not needs_image:
                continue

            slide_dimension = slide.get("content", {}).get("维度", dimension)

            user_message = self._load_llm_prompt(
                self.llm_prompt_file,
                dimension=slide_dimension,
                item=slide_title,
                product_domain=product_domain,
                image_base_path=os.path.join(os.path.dirname(__file__), "..", "..", "images")
            )

            try:
                llm_result = self.invoke_llm(user_message)
                self._safe_print(f"[ImageMatcher] LLM: {llm_result[:200]}...")

                result = self._parse_image_result(llm_result)
                if result.get("image_path"):
                    image_mapping[i] = result["image_path"]
            except Exception as e:
                self._safe_print(f"[ImageMatcher] LLM matching failed: {e}")

        return image_mapping

    def _rule_based_match_all(self, slides: List[Dict]) -> Dict:
        """基于规则为所有幻灯片匹配图片"""
        image_mapping = {}
        base_path = os.path.join(os.path.dirname(__file__), "..", "..", "images")

        for i, slide in enumerate(slides):
            slide_type = slide.get("type", "")
            if slide_type not in ["dimension_detail", "background", "intro"]:
                continue

            slide_title = slide.get("title", "")
            dimension = slide.get("content", {}).get("维度", "")

            image_path = self._rule_based_match(dimension, slide_title, "")
            if image_path:
                image_mapping[i] = image_path

        return image_mapping

    def _rule_based_match(self, dimension: str, item: str, product_domain: str) -> Optional[str]:
        """基于规则的图片匹配"""
        base_path = os.path.join(os.path.dirname(__file__), "..", "..", "images")

        if product_domain:
            exact_path = os.path.join(base_path, "zh", dimension, f"{item}_{product_domain}.png")
            if os.path.exists(exact_path):
                return exact_path

        fuzzy_path = os.path.join(base_path, "zh", dimension, f"{item}_默认.png")
        if os.path.exists(fuzzy_path):
            return fuzzy_path

        default_path = os.path.join(base_path, "zh", dimension, "默认.png")
        if os.path.exists(default_path):
            return default_path

        return None

    def _parse_image_result(self, llm_result: str) -> Dict:
        """解析LLM返回的图片路径"""
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
            self._safe_print(f"[ImageMatcher] Parse failed: {e}")

        return {"status": "not_found", "image_path": "", "match_type": "未匹配"}
