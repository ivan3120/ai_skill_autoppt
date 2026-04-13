#!/bin/bash
# AI Skill AutoPPT 安装脚本

set -e

echo "📦 正在安装 ai_skill_autoppt..."

# 检查Python版本
python_version=$(python --version 2>&1 | grep -oP '\d+\.\d+' | head -1)
required_version="3.8"

if [ $(echo "$python_version < $required_version" | bc) ]; then
    echo "❌ 需要Python 3.8或更高版本，当前版本: $python_version"
    exit 1
fi

echo "✅ Python版本检查通过: $python_version"

# 安装依赖
echo "📚 安装依赖包..."
pip install -e .

# 创建必要目录
echo "📁 创建必要目录..."
mkdir -p storage/ppt
mkdir -p storage/excel
mkdir -p storage/logs

# 复制配置示例
if [ ! -f config.yaml ]; then
    echo "📝 创建配置文件..."
    cp config.yaml config.yaml.example
fi

echo "✅ 安装完成!"
echo ""
echo "使用方法:"
echo "  ai-skill-autoppt generate --excel data.xlsx"
echo ""
echo "文档: https://github.com/example/ai-skill-autoppt"
