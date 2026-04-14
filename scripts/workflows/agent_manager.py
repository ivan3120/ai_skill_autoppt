# -*- coding: utf-8 -*-
"""
AutoPPT Agent模块 - 工作流编排

使用解耦的Agent模块进行多Agent协作。
"""

import os
import uuid
import concurrent.futures
from typing import Dict

# 导入配置
from scripts.config import Config

# 导入各个解耦的Agent
from scripts.agents.base import BaseAgent, AgentResult
from scripts.agents.data_analyst import DataAnalystAgent
from scripts.agents.content_planner import ContentPlannerAgent
from scripts.agents.visual_designer import VisualDesignerAgent
from scripts.agents.image_matcher import ImageMatcherAgent
from scripts.agents.quality_reviewer import QualityReviewerAgent
from scripts.agents.modification_handler import ModificationHandlerAgent


# ==================== Agent工厂 ====================

def get_agent(agent_type: str) -> BaseAgent:
    """获取Agent实例"""
    agents = {
        "data_analyst": DataAnalystAgent,
        "content_planner": ContentPlannerAgent,
        "visual_designer": VisualDesignerAgent,
        "image_matcher": ImageMatcherAgent,
        "quality_reviewer": QualityReviewerAgent,
        "modification_handler": ModificationHandlerAgent
    }

    agent_class = agents.get(agent_type)
    if agent_class:
        return agent_class()

    raise ValueError(f"未知的Agent类型: {agent_type}")


def is_llm_available() -> bool:
    """检查LLM是否可用（支持多种提供商）"""
    import os
    # 支持多种环境变量
    if (os.environ.get("ANTHROPIC_API_KEY") or
        os.environ.get("LLM_API_KEY") or
        os.environ.get("MINIMAX_API_KEY") or
        os.environ.get("OPENAI_API_KEY")):
        return True
    return Config.is_llm_available()


# ==================== 工作流编排 ====================

def generate_ppt_workflow(excel_path: str, user_selections: Dict) -> Dict:
    """
    执行PPT生成完整工作流
    """
    print("\n" + "="*50)
    print("AutoPPT 工作流开始")
    print("="*50)

    # Step 1: 数据解析 - DataAnalystAgent
    print("\n[Step 1] 数据解析...")
    data_analyst = get_agent("data_analyst")
    parse_result = data_analyst.execute({
        "excel_path": excel_path,
        "use_llm": user_selections.get("use_llm", False)
    })

    if parse_result.status != "success":
        return {
            "status": "error",
            "message": f"数据解析失败: {parse_result.message}"
        }

    parsed_data = parse_result.data
    print(f"  -> 解析完成: {parse_result.message}")

    # Step 2: 内容规划 - ContentPlannerAgent
    print("\n[Step 2] 内容规划...")
    content_planner = get_agent("content_planner")
    plan_result = content_planner.execute({
        "selected_dimensions": user_selections.get("dimensions", []),
        "selected_domains": user_selections.get("domains", []),
        "parsed_data": parsed_data
    })

    if plan_result.status != "success":
        return {
            "status": "error",
            "message": f"内容规划失败: {plan_result.message}"
        }

    slides = plan_result.data.get("slides", [])
    print(f"  -> 规划完成: {plan_result.message}")

    # Step 3 & 4: 并行执行风格学习和图片匹配
    print("\n[Step 3 & 4] 风格学习 + 图片匹配 (并行)...")

    visual_designer = get_agent("visual_designer")
    image_matcher = get_agent("image_matcher")

    # 构建模板路径 - 需要3层父目录才能到达项目根目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    scripts_dir = os.path.dirname(script_dir)
    skill_dir = os.path.dirname(scripts_dir)
    template_dir = os.path.join(skill_dir, "templates")

    template_path = user_selections.get("template", "")
    if template_path and not os.path.isabs(template_path):
        # 用户输入可能是 "default" -> "default_template.pptx" 或 "custom" -> "custom.pptx"
        if template_path == "default":
            template_path = "default_template"
        if not template_path.endswith(".pptx"):
            template_path = f"{template_path}.pptx"
        template_path = os.path.join(template_dir, template_path)
    elif not template_path:
        # 默认使用default_template.pptx
        template_path = os.path.join(template_dir, "default_template.pptx")

    print(f"  -> 使用模板: {template_path}")

    # 从parsed_data中提取有效参数
    selected_dimensions = user_selections.get("dimensions", [])
    selected_domains = user_selections.get("domains", [])
    dimension = selected_dimensions[0] if selected_dimensions else ""
    item = parsed_data.get("items", [{}])[0].get("name", "") if parsed_data.get("items") else ""
    product_domain = selected_domains[0] if selected_domains else ""

    # 并行执行
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
        future_style = executor.submit(visual_designer.execute, {
            "template_path": template_path if os.path.exists(template_path) else ""
        })
        future_image = executor.submit(image_matcher.execute, {
            "slides": slides,
            "dimension": dimension,
            "item": item,
            "product_domain": product_domain
        })

        style_result = future_style.result()
        image_result = future_image.result()

    style_config = style_result.data.get("style_config", {})
    image_mapping = image_result.data.get("image_mapping", {})
    print(f"  -> 风格学习完成: {style_result.message}")
    print(f"  -> 图片匹配完成: {image_result.message}")

    # Step 5: PPT生成
    print("\n[Step 5] PPT生成...")
    from scripts.services.ppt_generator import PPTGeneratorV2 as PPTGenerator
    import datetime

    # skill目录
    skill_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    base_dir = os.path.join(skill_dir, "storage")
    session_id_val = str(uuid.uuid4())

    # 传入template_path以提取模板风格
    generator = PPTGenerator(template_path=template_path)

    # 根据PPTGeneratorV2的实际接口构造参数
    content = {
        "slides": slides,
        "style_config": style_config,
        "image_mapping": image_mapping,
        "session_id": session_id_val
    }
    # create() 返回 Presentation 对象，需要调用 save() 保存
    generator.create(content=content)
    # 获取输出路径并保存
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    ppt_path = os.path.join(base_dir, "ppt", f"report_{session_id_val[:8]}_{timestamp}.pptx")
    # 调用save()将PPT保存到磁盘
    ppt_path = generator.save(ppt_path)

    print(f"  -> PPT生成完成: {ppt_path}")

    # Step 6: 质量审核 - QualityReviewerAgent
    print("\n[Step 6] 质量审核...")
    quality_reviewer = get_agent("quality_reviewer")
    review_result = quality_reviewer.execute({
        "ppt_content": slides,
        "original_data": parsed_data,
        "style_config": style_config
    })

    print(f"  -> 审核完成: {review_result.message}")

    print("\n" + "="*50)
    print("AutoPPT 工作流完成")
    print("="*50)

    return {
        "status": "success",
        "data": {
            "ppt_path": ppt_path,
            "slides_count": len(slides),
            "review_result": review_result.data
        },
        "message": f"PPT生成成功，共 {len(slides)} 页"
    }
