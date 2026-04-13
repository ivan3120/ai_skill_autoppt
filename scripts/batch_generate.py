# -*- coding: utf-8 -*-
"""
批量生成脚本

用于批量处理多个Excel文件生成PPT报告。
"""

import argparse
import os
import sys
from pathlib import Path
from datetime import datetime

# 添加scripts目录到路径
sys.path.insert(0, str(Path(__file__).parent))

from scripts import AutoPPTSkill


def batch_generate(excel_dir: str, output_dir: str, template: str = "default"):
    """
    批量生成PPT报告

    Args:
        excel_dir: Excel文件目录
        output_dir: 输出目录
        template: 模板名称
    """
    excel_path = Path(excel_dir)
    output_path = Path(output_dir)

    # 创建输出目录
    output_path.mkdir(parents=True, exist_ok=True)

    # 查找所有Excel文件
    excel_files = list(excel_path.glob("*.xlsx")) + list(excel_path.glob("*.xls"))

    if not excel_files:
        print(f"❌ 在 {excel_dir} 中没有找到Excel文件")
        return

    print(f"📁 找到 {len(excel_files)} 个Excel文件")
    print("")

    skill = AutoPPTSkill()
    success_count = 0
    fail_count = 0

    for excel_file in excel_files:
        filename = excel_file.stem
        output_file = output_path / f"{filename}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"

        print(f"📊 正在处理: {excel_file.name}")

        try:
            result = skill.generate(
                excel_path=str(excel_file),
                template=template,
                output_path=str(output_file)
            )

            if result.get("status") == "success":
                print(f"   ✅ 成功: {output_file.name}")
                success_count += 1
            else:
                print(f"   ❌ 失败: {result.get('message')}")
                fail_count += 1

        except Exception as e:
            print(f"   ❌ 异常: {str(e)}")
            fail_count += 1

        print("")

    print("=" * 50)
    print(f"✅ 完成! 成功: {success_count}, 失败: {fail_count}")
    print(f"📁 输出目录: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="批量生成PPT报告")
    parser.add_argument(
        "--excel-dir", "-i",
        required=True,
        help="Excel文件目录"
    )
    parser.add_argument(
        "--output-dir", "-o",
        default="./output",
        help="输出目录 (default: ./output)"
    )
    parser.add_argument(
        "--template", "-t",
        default="default",
        help="模板名称 (default: default)"
    )

    args = parser.parse_args()

    batch_generate(args.excel_dir, args.output_dir, args.template)


if __name__ == "__main__":
    main()
