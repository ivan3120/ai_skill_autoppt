# -*- coding: utf-8 -*-
"""
ContentPlanner Agent - 内容规划师

独立解耦的Agent模块，负责规划PPT报告的内容结构和故事线。
支持配置化幻灯片模板
"""

import json
import re
from typing import Dict, List
from scripts.agents.base import BaseAgent, AgentResult
from scripts.services.slide_config import SlideConfig, config as slide_config


class ContentPlannerAgent(BaseAgent):
    """
    内容规划师Agent

    角色：专业PPT内容规划师，专门规划云核心网网络架构评估报告的故事线和内容结构
    任务：基于解析后的本体数据，规划PPT报告的内容结构和故事线
    """

    def __init__(self):
        super().__init__("ContentPlanner", "内容规划师", "content_planner")
        # 加载幻灯片配置
        self.slide_config = slide_config

    def execute(self, input_data: Dict) -> AgentResult:
        """
        规划PPT内容结构
        """
        selected_dimensions = input_data.get("selected_dimensions", [])
        selected_domains = input_data.get("selected_domains", [])
        parsed_data = input_data.get("parsed_data", {})

        # 使用LLM进行智能规划
        if self._has_llm():
            # 从prompt文件加载LLM指令
            user_message = self._load_llm_prompt(
                self.llm_prompt_file,
                dimensions=', '.join(selected_dimensions),
                product_domains=parsed_data.get('产品域列表', []),
                overview_results=parsed_data.get('整体评估概览', [])
            )

            llm_result = self.invoke_llm(user_message, context=parsed_data)
            self._safe_print(f"[ContentPlanner] LLM: {llm_result[:200]}...")

            # 解析LLM结果生成幻灯片结构
            slides = self._parse_llm_plan(llm_result, selected_dimensions, parsed_data)
        else:
            # 规则引擎fallback
            slides = self._rule_based_plan(selected_dimensions, selected_domains, parsed_data)

        return AgentResult(
            status="success",
            data={"slides": slides},
            message=f"规划了{len(slides)}页内容"
        )

    def _rule_based_plan(self, dimensions: List, domains: List, data: Dict) -> List[Dict]:
        """
        基于规则的内容规划 - 使用灵活的列名匹配
        """
        slides = []
        overview_results = data.get("整体评估概览", [])
        product_domains = data.get("产品域列表", [])

        # 灵活的列名映射 - 尝试多种可能的列名
        def get_field(item, *field_names):
            """尝试获取任意可能的字段名"""
            for fn in field_names:
                if fn in item:
                    return item[fn]
                # 也尝试模糊匹配
                for k in item.keys():
                    if fn in k or k in fn:
                        return item[k]
            return None

        # 解析整体评估概览数据
        dimension_summary = {}
        for item in overview_results:
            dim_code = get_field(item, "评估维度", "维度", "维度名称") or ""
            if not dim_code:
                continue
            mt_code = dim_code.split("-")[0] if "-" in dim_code else dim_code

            if mt_code not in dimension_summary:
                dimension_summary[mt_code] = {
                    "维度": dim_code, "通过数": 0, "待改进数": 0, "不通过数": 0,
                    "通过率": "0%", "风险": "", "整体评估结论": "",
                    "主要风险描述": "", "优化建议": ""
                }

            summary = dimension_summary[mt_code]
            # 使用灵活的字段获取，使用英文变量名避免编码问题
            passed = get_field(item, "passed", "pass", "ok") or 0
            tofix = get_field(item, "tofix", "fix", "pending") or 0
            failed = get_field(item, "failed", "fail", "ng") or 0

            try:
                summary["passed"] = summary.get("passed", 0) + int(passed) if passed else 0
            except:
                pass
            try:
                summary["tofix"] = summary.get("tofix", 0) + int(tofix) if tofix else 0
            except:
                pass
            try:
                summary["failed"] = summary.get("failed", 0) + int(failed) if failed else 0
            except:
                pass

            total = summary["通过数"] + summary["待改进数"] + summary["不通过数"]
            if total > 0:
                summary["通过率"] = f"{int(summary['通过数'] * 100 / total)}%"
            if get_field(item, "风险等级", "风险", "等级"):
                summary["风险"] = get_field(item, "风险等级", "风险", "等级")
            if get_field(item, "整体评估结论", "评估结论", "结论"):
                summary["整体评估结论"] = get_field(item, "整体评估结论", "评估结论", "结论")
            if get_field(item, "主要风险描述", "风险描述", "描述"):
                summary["主要风险描述"] = get_field(item, "主要风险描述", "风险描述", "描述")
            if get_field(item, "优化建议", "建议", "suggestion"):
                summary["优化建议"] = get_field(item, "优化建议", "建议", "suggestion")

        # 生成highlights
        highlights = []
        for mt in ["MT0", "MT1", "MT2", "MT3", "MT4", "MT5", "MT6"]:
            if mt in dimension_summary:
                s = dimension_summary[mt]
                status = "需改进"
                if s["风险"] == "低":
                    status = "良好"
                highlights.append({
                    "维度": s["维度"], "通过率": s["通过率"], "风险": s["风险"], "状态": status
                })

        total_pass = sum(s["通过数"] for s in dimension_summary.values())
        total_all = total_pass + sum(s["待改进数"] + s["不通过数"] for s in dimension_summary.values())
        overall_rate = f"{int(total_pass * 100 / total_all)}%" if total_all > 0 else "0%"

        # 1. 封面页 - 使用配置
        cover_config = self.slide_config.get_slide_type('cover')
        cover_title = "云核心网网络架构评估报告"
        if domains:
            cover_title = f"{domains[0]}网络架构评估报告"

        slides.append({
            "type": "cover",
            "title": cover_title,
            "subtitle": cover_config.get("subtitle", "运营商客户汇报材料"),
            "date": cover_config.get("date", "2026年4月")
        })

        # 2. 目录页 - 使用配置
        toc_config = self.slide_config.get_slide_type('toc')
        toc_items = self.slide_config.build_toc_items(dimensions)
        slides.append({
            "type": "toc",
            "title": toc_config.get("title", "目录"),
            "items": toc_items
        })

        # 3. 评估方案介绍 - 使用配置
        intro_config = self.slide_config.get_slide_type('intro')
        slides.append({
            "type": "intro",
            "title": intro_config.get("title", "一、评估方案介绍"),
            "content": {
                "评估维度": dimensions,
                "评估产品域": domains if domains else [pd.get("name", "") for pd in product_domains[:5]],
                "评估方法": intro_config.get("default", {}).get("评估方法", "现场调研+系统检测+数据分析"),
                "评估标准": intro_config.get("default", {}).get("评估标准", "依据MT0-MT6七维评估体系")
            }
        })

        # 4. 业务背景 - 使用配置
        bg_config = self.slide_config.get_slide_type('background')
        domain_summary = []
        for pd in product_domains[:5]:
            pd_name = pd.get("name", "")
            if pd_name:
                domain_summary.append({
                    "name": pd_name, "ne_count": pd.get("ne_count", 0), "dr_count": pd.get("dr_count", 0)
                })

        total_ne = sum(d.get("ne_count", 0) for d in domain_summary)
        total_dr = sum(d.get("dr_count", 0) for d in domain_summary)

        slides.append({
            "type": "background",
            "title": bg_config.get("title", "二、业务背景与网络概况"),
            "content": {
                "总体网元数": total_ne if total_ne > 0 else len(overview_results) * 5,
                "容灾组数": total_dr if total_dr > 0 else len(product_domains) * 3,
                "产品域分布": domain_summary if domain_summary else [{"name": d, "ne_count": 10, "dr_count": 2} for d in domains[:4]]
            }
        })

        # 5. 整体评估概览 - 使用配置
        overview_config = self.slide_config.get_slide_type('overview')
        slides.append({
            "type": "overview",
            "title": overview_config.get("title", "三、整体评估概览"),
            "content": {
                "description": "各评估维度整体情况汇总",
                "highlights": highlights if highlights else [
                    {"维度": "MT0-网元高稳", "通过率": "80%", "风险": "中", "状态": "良好"},
                    {"维度": "MT1-部署离散度", "通过率": "65%", "风险": "高", "状态": "需改进"}
                ],
                "overall": f"整体评估结果: 良好（{overall_rate}通过率），部分维度存在优化空间"
            }
        })

        # 6-12. 各评估维度详情页 - 使用配置
        for idx, mt in enumerate(["MT0", "MT1", "MT2", "MT3", "MT4", "MT5", "MT6"]):
            if mt not in dimensions:
                continue

            dim_config = self.slide_config.get_dimension(mt)

            mt_items = [item for item in overview_results if item.get("评估维度", "").startswith(mt)]
            findings, risks, suggestions, pass_rates = [], [], [], []

            for item in mt_items:
                eval_item = item.get("评估项", "")
                pass_rate = item.get("通过率", "0%")
                pass_rates.append(pass_rate)
                findings.append(f"{eval_item}: {pass_rate}通过")
                if item.get("主要风险描述"):
                    risks.append(item.get("主要风险描述", ""))
                if item.get("优化建议"):
                    suggestions.append(item.get("优化建议", ""))

            avg_rate = pass_rates[0] if pass_rates else "0%"
            mt_summary = dimension_summary.get(mt, {})
            risk = mt_summary.get("风险", "低")
            if not risk:
                risk = "低"
                if any("高" in str(r) for r in risks):
                    risk = "高"
                elif any("中" in str(r) for r in risks):
                    risk = "中"

            # 使用配置获取标题和状态
            dim_title = self.slide_config.get_dimension_title(mt, dimensions)
            status = self.slide_config.get_status_for_risk(risk)

            slides.append({
                "type": "dimension_detail",
                "title": dim_title,
                "pass_rate": avg_rate,
                "risk": risk,
                "status": status,
                "overview": f"评估{dim_config.get('name', mt)}情况",
                "findings": findings[:3] if findings else ["数据评估中..."],
                "risks": risks[:2] if risks else ["暂无风险描述"],
                "suggestions": suggestions[:2] if suggestions else ["暂无优化建议"]
            })

        # 13. 总结与优化建议 - 使用配置
        summary_config = self.slide_config.get_slide_type('summary')
        high_risk_dims = [h["维度"] for h in highlights if h.get("风险") == "高"]
        summary_text = f"整体评估结果: {overall_rate}通过率"
        if high_risk_dims:
            summary_text += f"，{', '.join(high_risk_dims)}存在较高风险，需重点关注"

        slides.append({
            "type": "summary",
            "title": summary_config.get("title", "十一、总结与优化建议"),
            "content": {
                "overall": summary_text,
                "priorities": [
                    {"优先级": "高", "问题": dim, "建议": dimension_summary.get(dim.split('-')[0], {}).get("优化建议", "需优化")}
                    for dim in high_risk_dims[:3]
                ] if high_risk_dims else [{"优先级": "中", "问题": "整体评估良好", "建议": "继续保持"}],
                "next_steps": "建议分阶段实施优化，优先处理高风险项"
            }
        })

        return slides

    def _parse_llm_plan(self, llm_result: str, dimensions: List, data: Dict, domains: List = None) -> List[Dict]:
        """解析LLM的规划结果"""
        slides = []
        try:
            if "```json" in llm_result:
                start = llm_result.find("```json") + 7
                end = llm_result.find("```", start)
                json_str = llm_result[start:end].strip()
                raw_data = json.loads(json_str)

                raw_slides = raw_data
                if isinstance(raw_data, dict):
                    for key in ['slides', 'ppt_structure', 'pages']:
                        if key in raw_data and isinstance(raw_data[key], list):
                            raw_slides = raw_data[key]
                            break

                if isinstance(raw_slides, list) and len(raw_slides) > 0:
                    slides = self._convert_llm_slides(raw_slides, data)
        except Exception as e:
            pass

        if not slides:
            slides = self._rule_based_plan(dimensions, domains if domains else [], data)
            unique_slides = []
            seen_titles = set()
            for slide in slides:
                title = slide.get('title', '')
                if title not in seen_titles:
                    unique_slides.append(slide)
                    seen_titles.add(title)
            return unique_slides

        return slides

    def _convert_llm_slides(self, raw_slides, data: Dict) -> List[Dict]:
        """将LLM输出的幻灯片格式转换为ppt_generator期望的格式"""
        converted = []

        if not isinstance(raw_slides, list):
            return []

        for slide in raw_slides:
            if not isinstance(slide, dict):
                continue

            slide_type = slide.get("type") or slide.get("slide_type") or ""
            title = slide.get("title") or slide.get("page_title") or ""

            if not slide_type:
                if "封面" in title:
                    slide_type = "cover"
                elif "目录" in title:
                    slide_type = "toc"
                elif "评估方案" in title or "方案介绍" in title:
                    slide_type = "intro"
                elif "背景" in title or "概况" in title:
                    slide_type = "background"
                elif "整体" in title or "概览" in title or "结果" in title:
                    slide_type = "overview"
                elif "总结" in title or "建议" in title:
                    slide_type = "summary"
                elif "MT" in title or "评估" in title:
                    slide_type = "dimension_detail"

            main_content = slide.get("main_content", [])
            converted_slide = {"type": slide_type, "title": title}

            if slide_type == "cover":
                converted_slide["subtitle"] = slide.get("subtitle", "运营商客户汇报材料")
                converted_slide["date"] = slide.get("date", "2026年4月")
                if main_content:
                    converted_slide["subtitle"] = main_content[0].replace("报告标题：", "") if main_content else "运营商客户汇报材料"

            elif slide_type in ["intro", "toc"]:
                converted_slide["content"] = {
                    "评估维度": "MT0, MT1, MT2, MT3, MT4, MT5, MT6",
                    "评估产品域": [pd.get("name", "") for pd in data.get("产品域列表", [])],
                    "评估方法": "现场调研+系统检测+数据分析",
                    "评估标准": "依据MT0-MT6七维评估体系"
                }

            elif slide_type == "background":
                total_ne = 0
                total_dr = 0
                for line in main_content:
                    if "网元" in line and any(c.isdigit() for c in line):
                        nums = re.findall(r'\d+', line)
                        if nums:
                            total_ne = int(nums[0])
                    if "容灾" in line and any(c.isdigit() for c in line):
                        nums = re.findall(r'\d+', line)
                        if nums:
                            total_dr = int(nums[0])

                converted_slide["content"] = {
                    "总体网元数": total_ne or sum(pd.get("ne_count", 0) for pd in data.get("产品域列表", [])),
                    "容灾组数": total_dr or sum(pd.get("dr_count", 0) for pd in data.get("产品域列表", [])),
                    "产品域分布": [{"name": pd.get("name", ""), "ne_count": pd.get("ne_count", 0), "dr_count": pd.get("dr_count", 0)} for pd in data.get("产品域列表", [])]
                }

            elif slide_type == "overview":
                highlights = []
                for line in main_content:
                    if "MT" in line and ("通过率" in line or "%" in line):
                        mt_match = re.search(r'(MT\d)', line)
                        rate_match = re.search(r'(\d+%)', line)
                        risk_match = re.search(r'(高|中|低)', line)
                        if mt_match:
                            highlights.append({
                                "维度": mt_match.group(1),
                                "通过率": rate_match.group(1) if rate_match else "",
                                "风险": risk_match.group(1) if risk_match else "中",
                                "状态": "需改进" if risk_match and risk_match.group(1) == "高" else "良好"
                            })

                if not highlights:
                    overview = data.get("整体评估概览", [])
                    mt_codes = ["MT0", "MT1", "MT2", "MT3", "MT4", "MT5", "MT6"]
                    for mt in mt_codes:
                        mt_items = [o for o in overview if o.get("评估维度", "").startswith(mt)]
                        if mt_items:
                            passed = sum(int(o.get("通过数", 0) or 0) for o in mt_items)
                            total = passed + sum(int(o.get("待改进数", 0) or 0) for o in mt_items) + sum(int(o.get("不通过数", 0) or 0) for o in mt_items)
                            rate = f"{int(passed * 100 / total)}%" if total > 0 else "0%"
                            risk = mt_items[0].get("风险等级", "中")
                            highlights.append({
                                "维度": mt, "通过率": rate, "风险": risk,
                                "状态": "需改进" if risk == "高" else "良好"
                            })

                converted_slide["content"] = {"整体评估结果": "整体评估完成", "highlights": highlights}

            elif slide_type == "dimension_detail":
                mt_match = re.search(r'(MT\d)', title)
                mt_code = mt_match.group(1) if mt_match else ""

                overview = data.get("整体评估概览", [])
                mt_items = [o for o in overview if o.get("评估维度", "").startswith(mt_code)] if mt_code else []

                findings, risks, suggestions = [], [], []
                for item in mt_items:
                    findings.append(f"{item.get('评估项', '')}: {item.get('通过率', '')}通过")
                    if item.get("主要风险描述"):
                        risks.append(item.get("主要风险描述", ""))
                    if item.get("优化建议"):
                        suggestions.append(item.get("优化建议", ""))

                converted_slide["content"] = {
                    "findings": findings[:3], "risks": risks[:2], "suggestions": suggestions[:2]
                }

            elif slide_type == "summary":
                converted_slide["content"] = {
                    "overall": slide.get("summary", "整体评估完成"),
                    "priorities": slide.get("priorities", []),
                    "next_steps": slide.get("next_steps", "")
                }

            converted.append(converted_slide)

        return converted
