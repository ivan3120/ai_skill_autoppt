# -*- coding: utf-8 -*-
"""
Base Agent - Agent基类定义

包含BaseAgent和AgentResult两个基础类。
设计原则：SKILL无LLM配置依赖，规则引擎内置，宿主Agent可选择注入LLM增强能力。
"""

import os
import json
from typing import Dict, Any, List, Optional, Union
from dataclasses import dataclass


@dataclass
class AgentResult:
    """Agent执行结果"""
    status: str  # success/error/partial
    data: Any = None
    message: str = ""


class BaseAgent:
    """Agent基类 - 定义Agent接口规范

    设计原则：
    - SKILL无LLM配置依赖，内置规则引擎可独立运行
    - 宿主Agent可选择注入LLM客户端以增强能力
    - 无LLM注入时，所有Agent使用内置规则引擎工作
    """

    # 类级别的LLM客户端（可被注入）
    _llm_client: Optional[Any] = None
    _llm_config: Dict = {}

    def __init__(self, name: str, role: str, skill_id: str = None, llm_client: Any = None):
        """
        初始化Agent

        Args:
            name: Agent名称
            role: Agent角色描述
            skill_id: Skill ID（从Skill注册中心加载）
            llm_client: (可选) 宿主Agent注入的LLM客户端，不注入则使用规则引擎
        """
        self.name = name
        self.role = role
        self.skill_id = skill_id
        self.prompt_file = None
        self.prompt_template = ""
        self.llm = llm_client  # 优先使用注入的客户端
        self._client = None
        self.llm_prompt_file = None  # LLM指令文件（用于运行时加载）

        # 优先从Skill注册中心加载
        if skill_id:
            self._load_from_skill(skill_id)
        # 兼容旧方式：直接加载prompt文件
        elif role:
            # 尝试从skills目录加载同名的skill
            self._load_from_skill(self._get_skill_id_from_role(role))

    def _get_skill_id_from_role(self, role: str) -> str:
        """从role名称推断skill_id"""
        role_to_skill = {
            "数据分析师": "data_analyst",
            "内容规划师": "content_planner",
            "视觉设计师": "visual_designer",
            "图片匹配师": "image_matcher",
            "质量审核员": "quality_reviewer",
            "修改处理器": "modification_handler"
        }
        return role_to_skill.get(role, "")

    def _load_from_skill(self, skill_id: str):
        """从Skill注册中心加载"""
        if not skill_id:
            return

        try:
            # Skill注册中心已移除，跳过
            self._safe_print(f"[Agent] Skill registry not available: {skill_id}")
        except Exception as e:
            self._safe_print(f"[Agent] Failed to load skill {skill_id}: {e}")
            self._safe_print(f"[Agent] Failed to load skill {skill_id}: {e}")

    @classmethod
    def set_llm_client(cls, client: Any):
        """设置全局LLM客户端（供宿主Agent调用）"""
        cls._llm_client = client

    @classmethod
    def set_llm_config(cls, config: Dict):
        """设置全局LLM配置（供宿主Agent调用）"""
        cls._llm_config = config

    @classmethod
    def get_llm_client(cls) -> Optional[Any]:
        """获取LLM客户端（优先使用注入的，否则使用类级别的）"""
        if cls._llm_client:
            return cls._llm_client
        return None

    def _init_llm(self):
        """初始化LLM客户端 - 不再从Config读取，由宿主Agent注入"""
        # 移除Config依赖，SKILL不主动初始化LLM
        # 宿主Agent可通过set_llm_client()注入
        pass

    def _safe_print(self, msg: str, limit: int = 200):
        """Safe print to avoid encoding errors"""
        try:
            print(msg[:limit] + "..." if len(msg) > limit else msg)
        except UnicodeEncodeError:
            print("[LLM Response - chars truncated]")

    def _load_llm_prompt(self, skill_id: str = None, **kwargs) -> str:
        """
        加载LLM prompt - 优先从Skill注册中心加载

        Args:
            skill_id: Skill ID（可选，默认使用self.llm_prompt_file）
            **kwargs: 要替换的占位符
        """
        # 优先从文件方式加载
        use_skill_id = skill_id or self.llm_prompt_file

        if not use_skill_id:
            return ""

        # 从文件方式加载
        if os.path.exists(use_skill_id):
            with open(use_skill_id, 'r', encoding='utf-8') as f:
                content = f.read()
            for key, value in kwargs.items():
                placeholder = f"{{{key}}}"
                content = content.replace(placeholder, str(value))
            return content

        return ""

    def execute(self, input_data: Dict) -> AgentResult:
        """
        执行Agent任务 - 由子类实现具体逻辑

        Args:
            input_data: 输入数据字典

        Returns:
            AgentResult: 执行结果
        """
        raise NotImplementedError

    def invoke_llm(self, user_message: str, context: Dict = None) -> str:
        """
        调用LLM获取自然语言响应

        Args:
            user_message: 用户消息
            context: 上下文数据

        Returns:
            LLM响应文本
        """
        # 无LLM时返回基于规则的响应
        if not self._has_llm():
            return self._fallback_response(user_message, context)

        # 构建系统消息
        system_message = self.prompt_template
        if context:
            context_json = json.dumps(context, ensure_ascii=False, indent=2)
            system_message += f"\n\n## 上下文数据\n```json\n{context_json}\n```"

        # 优先使用实例级LLM客户端
        llm_client = self.llm or BaseAgent._llm_client

        # 使用LangChain客户端（通用接口）
        if llm_client:
            return self._invoke_langchain(llm_client, system_message, user_message)

        # 使用直接SDK
        if self._client:
            return self._invoke_direct_sdk(system_message, user_message)

        return self._fallback_response(user_message, context)

    def _has_llm(self) -> bool:
        """检查是否有可用的LLM（实例级或类级别）"""
        return self.llm is not None or self._client is not None or BaseAgent._llm_client is not None

    def _invoke_langchain(self, llm_client, system_message: str, user_message: str) -> str:
        """使用通用LLM客户端调用LLM（支持LangChain/OpenAI等）"""
        try:
            # 尝试统一接口：传入消息列表
            from langchain_core.messages import HumanMessage, SystemMessage
            messages = [
                SystemMessage(content=system_message),
                HumanMessage(content=user_message)
            ]
            response = llm_client.invoke(messages)

            # 处理返回格式 - 可能是字符串或对象
            if hasattr(response, 'content'):
                content = response.content
                # 处理MiniMax返回的格式 - 可能是list包含thinking和text
                if isinstance(content, list):
                    for item in content:
                        if isinstance(item, dict) and item.get('type') == 'text':
                            return item.get('text', '')
                    return str(content[0]) if content else ''
                return content
            return str(response)
        except ImportError:
            # 如果没有langchain，尝试直接调用
            try:
                return llm_client.invoke(system_message + "\n\n" + user_message)
            except Exception as e:
                return f"LLM调用失败: {str(e)}"
        except Exception as e:
            return f"LLM调用失败: {str(e)}"

    def _invoke_direct_sdk(self, system_message: str, user_message: str) -> str:
        """直接使用Anthropic SDK调用LLM"""
        try:
            llm_config = Config.get_llm_config()

            # 构建完整消息
            full_message = f"{system_message}\n\n用户请求: {user_message}"

            message = self._client.messages.create(
                model=llm_config.model,
                max_tokens=llm_config.max_tokens,
                temperature=llm_config.temperature,
                messages=[{'role': 'user', 'content': full_message}]
            )

            # 解析响应
            for content in message.content:
                if hasattr(content, 'text'):
                    return content.text

            return "未获取到有效响应"
        except Exception as e:
            return f"LLM调用失败: {str(e)}"

    def _fallback_response(self, user_message: str, context: Dict = None) -> str:
        """无LLM时的降级响应"""
        return "LLM未配置，使用规则引擎响应"

    def parse_llm_json_result(self, llm_result: str, default: Dict = None) -> Dict:
        """
        统一的LLM JSON解析方法 - 消除重复代码

        Args:
            llm_result: LLM返回的字符串
            default: 解析失败时返回的默认值

        Returns:
            Dict: 解析后的JSON对象
        """
        default = default or {}
        if not llm_result:
            return default

        try:
            # 尝试从markdown代码块中提取
            if "```json" in llm_result:
                start = llm_result.find("```json") + 7
                end = llm_result.find("```", start)
                if end > start:
                    return json.loads(llm_result[start:end].strip())
            elif "```" in llm_result:
                start = llm_result.find("```") + 3
                end = llm_result.find("```", start)
                if end > start:
                    return json.loads(llm_result[start:end].strip())
            # 尝试直接查找JSON对象
            elif "{" in llm_result:
                start = llm_result.find("{")
                end = llm_result.rfind("}") + 1
                if end > start:
                    return json.loads(llm_result[start:end])
        except Exception as e:
            self._safe_print(f"[LLM Parse] Failed: {e}")
        return default

    def execute_with_fallback(self, input_data: Dict, llm_handler, rule_handler, required_keys: List[str] = None) -> AgentResult:
        """
        统一的execute模板：LLM优先，规则降级

        Args:
            input_data: 输入数据字典
            llm_handler: LLM处理函数
            rule_handler: 规则处理函数
            required_keys: 必需的键列表

        Returns:
            AgentResult: 执行结果
        """
        # 参数校验
        if required_keys:
            missing = [k for k in required_keys if k not in input_data]
            if missing:
                return AgentResult(status="error", message=f"Missing required keys: {missing}")

        if self._has_llm():
            try:
                result = llm_handler(input_data)
                return AgentResult(status="success", data=result, message="")
            except Exception as e:
                self._safe_print(f"[LLM] Failed, using fallback: {e}")

        return AgentResult(status="success", data=rule_handler(input_data), message="(规则引擎)")
