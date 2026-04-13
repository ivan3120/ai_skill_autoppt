# -*- coding: utf-8 -*-
"""
CLI - 命令行入口
"""

import argparse
import sys
import os
from pathlib import Path

# 添加scripts目录到路径
sys.path.insert(0, str(Path(__file__).parent))


def main():
    """主入口函数"""
    parser = argparse.ArgumentParser(
        description="AI Skill: 云核心网评估报告生成",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  ai-skill-autoppt generate --excel data.xlsx --output report.pptx
  ai-skill-autoppt generate --excel data.xlsx --template custom --output report.pptx
  ai-skill-autoppt list-templates
  ai-skill-autoppt --version
        """
    )

    subparsers = parser.add_subparsers(dest="command", help="可用命令")

    # generate命令
    generate_parser = subparsers.add_parser("generate", help="生成PPT报告")
    generate_parser.add_argument(
        "--excel", "-e",
        required=True,
        help="Excel评估数据文件路径"
    )
    generate_parser.add_argument(
        "--output", "-o",
        default="report.pptx",
        help="输出PPT文件路径 (default: report.pptx)"
    )
    generate_parser.add_argument(
        "--template", "-t",
        default="default",
        help="PPT模板名称 (default: default)"
    )
    generate_parser.add_argument(
        "--dimensions",
        nargs="+",
        default=["MT0", "MT1", "MT2", "MT3", "MT4", "MT5", "MT6"],
        help="评估维度 (default: MT0-MT6)"
    )
    generate_parser.add_argument(
        "--use-llm",
        action="store_true",
        help="使用LLM增强"
    )

    # list-templates命令
    subparsers.add_parser("list-templates", help="列出可用模板")

    # parse命令 - 仅解析数据不生成PPT
    parse_parser = subparsers.add_parser("parse", help="解析Excel数据")
    parse_parser.add_argument(
        "--excel", "-e",
        required=True,
        help="Excel文件路径"
    )
    parse_parser.add_argument(
        "--output", "-o",
        help="输出JSON文件路径 (可选)"
    )

    # version
    parser.add_argument(
        "--version", "-v",
        action="version",
        version="ai-skill-autoppt v1.0.0"
    )

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return

    if args.command == "generate":
        generate_report(args)
    elif args.command == "list-templates":
        list_templates()
    elif args.command == "parse":
        parse_excel(args)


def generate_report(args):
    """生成PPT报告"""
    from scripts.workflows.agent_manager import generate_ppt_workflow

    user_selections = {
        "dimensions": args.dimensions,
        "domains": [],
        "template": args.template,
        "use_llm": args.use_llm
    }

    print(f"📊 正在解析Excel文件: {args.excel}")
    print(f"📝 使用模板: {args.template}")
    print(f"📐 评估维度: {', '.join(args.dimensions)}")

    result = generate_ppt_workflow(args.excel, user_selections)

    if result.get("status") == "success":
        ppt_path = result.get("data", {}).get("ppt_path")
        slides_count = result.get("data", {}).get("slides_count")

        # 移动到指定输出路径
        if args.output and args.output != ppt_path:
            import shutil
            shutil.copy(ppt_path, args.output)
            ppt_path = args.output

        print(f"✅ PPT生成成功!")
        print(f"   路径: {ppt_path}")
        print(f"   页数: {slides_count}")
    else:
        print(f"❌ 生成失败: {result.get('message')}")
        sys.exit(1)


def list_templates():
    """列出可用模板"""
    from pathlib import Path

    template_dir = Path(__file__).parent.parent / "resources" / "templates"

    if not template_dir.exists():
        print("❌ 模板目录不存在")
        return

    templates = list(template_dir.glob("*.pptx"))

    if not templates:
        print("📁 没有找到模板文件")
        return

    print("📁 可用模板:")
    for t in templates:
        print(f"   - {t.stem}")


def parse_excel(args):
    """解析Excel数据"""
    from scripts.services.excel_parser import ExcelParser
    import json

    print(f"📊 正在解析: {args.excel}")

    try:
        parser = ExcelParser(args.excel)
        parser.load()
        ontology = parser.parse_all()

        result = ontology.to_dict()

        if args.output:
            with open(args.output, "w", encoding="utf-8") as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            print(f"✅ 数据已保存到: {args.output}")
        else:
            print(json.dumps(result, ensure_ascii=False, indent=2))

    except Exception as e:
        print(f"❌ 解析失败: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
