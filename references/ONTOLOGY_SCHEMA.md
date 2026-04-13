# 云核心网评估报告 - 本体模型JSON Schema

## 概述

本文档定义云核心网网络架构评估的标准化本体模型JSON Schema，用于规范评估数据的表示和交换。

## 完整JSON Schema

```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "title": "CloudCoreNetworkEvaluation",
  "description": "云核心网网络架构评估本体模型",
  "type": "object",
  "required": ["metadata", "product_domains", "evaluation_dimensions", "evaluation_items"],
  "properties": {
    "metadata": {
      "type": "object",
      "description": "元数据",
      "required": ["source_file", "parse_time", "version"],
      "properties": {
        "source_file": {
          "type": "string",
          "description": "源Excel文件路径"
        },
        "parse_time": {
          "type": "string",
          "format": "date-time",
          "description": "解析时间"
        },
        "version": {
          "type": "string",
          "description": "Schema版本",
          "default": "1.0.0"
        },
        "operator": {
          "type": "string",
          "description": "运营商名称"
        },
        "report_date": {
          "type": "string",
          "description": "报告日期"
        }
      }
    },
    "product_domains": {
      "type": "array",
      "description": "产品域列表",
      "items": {
        "type": "object",
        "required": ["name"],
        "properties": {
          "name": {
            "type": "string",
            "description": "产品域名称"
          },
          "ne_count": {
            "type": "integer",
            "description": "网元数量"
          },
          "dr_count": {
            "type": "integer",
            "description": "容灾组数量"
          }
        }
      }
    },
    "network_elements": {
      "type": "array",
      "description": "网元列表",
      "items": {
        "type": "object",
        "required": ["name", "product_domain"],
        "properties": {
          "name": {
            "type": "string",
            "description": "网元名称"
          },
          "product_domain": {
            "type": "string",
            "description": "所属产品域"
          },
          "disaster_recovery_group": {
            "type": "string",
            "description": "容灾组"
          },
          "status": {
            "type": "string",
            "enum": ["normal", "warning", "critical"],
            "description": "状态"
          },
          "dc_location": {
            "type": "string",
            "description": "所在数据中心"
          }
        }
      }
    },
    "evaluation_dimensions": {
      "type": "array",
      "description": "评估维度定义",
      "items": {
        "type": "object",
        "required": ["code", "name"],
        "properties": {
          "code": {
            "type": "string",
            "description": "维度代码 (如 MT0, MT1)"
          },
          "name": {
            "type": "string",
            "description": "维度名称"
          },
          "description": {
            "type": "string",
            "description": "维度说明"
          },
          "weight": {
            "type": "number",
            "description": "权重"
          }
        }
      }
    },
    "evaluation_items": {
      "type": "array",
      "description": "评估项列表",
      "items": {
        "type": "object",
        "required": ["dimension", "name", "score", "pass_rate"],
        "properties": {
          "dimension": {
            "type": "string",
            "description": "所属维度代码"
          },
          "name": {
            "type": "string",
            "description": "评估项名称"
          },
          "score": {
            "type": "integer",
            "description": "总得分"
          },
          "pass_rate": {
            "type": "string",
            "description": "通过率"
          },
          "level": {
            "type": "string",
            "enum": ["优", "良", "中", "差"],
            "description": "通过等级"
          },
          "result": {
            "type": "string",
            "enum": ["通过", "待改进", "不通过"],
            "description": "评估结果"
          },
          "findings": {
            "type": "string",
            "description": "主要发现"
          },
          "suggestions": {
            "type": "string",
            "description": "优化建议"
          }
        }
      }
    },
    "overview": {
      "type": "object",
      "description": "整体评估概览",
      "properties": {
        "total_score": {
          "type": "integer",
          "description": "总体得分"
        },
        "overall_pass_rate": {
          "type": "string",
          "description": "整体通过率"
        },
        "risk_level": {
          "type": "string",
          "enum": ["低", "中", "高"],
          "description": "整体风险等级"
        },
        "summary": {
          "type": "string",
          "description": "整体评估总结"
        }
      }
    }
  }
}
```

## 示例数据

```json
{
  "metadata": {
    "source_file": "assets/test_operator_A_5G.xlsx",
    "parse_time": "2026-04-08T18:58:31.581950",
    "version": "1.0.0",
    "operator": "运营商A",
    "report_date": "2026-04"
  },
  "product_domains": [
    {"name": "5G核心网", "ne_count": 15, "dr_count": 8},
    {"name": "物联网专用网元", "ne_count": 15, "dr_count": 8},
    {"name": "VoNR核心网", "ne_count": 15, "dr_count": 8},
    {"name": "短信中心", "ne_count": 15, "dr_count": 8}
  ],
  "network_elements": [
    {"name": "5G核心网-A1", "product_domain": "5G核心网", "disaster_recovery_group": "", "status": "normal"},
    {"name": "5G核心网-B1", "product_domain": "5G核心网", "disaster_recovery_group": "", "status": "normal"}
  ],
  "evaluation_dimensions": [
    {"code": "MT0", "name": "网元高稳评估", "description": "评估网元的高可用性和稳定性"},
    {"code": "MT1", "name": "部署离散度评估", "description": "评估网元在不同站点的分布情况"},
    {"code": "MT2", "name": "组网架构高可用评估", "description": "评估网络的冗余设计"},
    {"code": "MT3", "name": "业务路由高可用评估", "description": "评估业务路由的可靠性"},
    {"code": "MT4", "name": "网络容量均衡度评估", "description": "评估网元间的负载均衡"},
    {"code": "MT5", "name": "网元版本EOS评估", "description": "评估网元版本的EOS状态"},
    {"code": "MT6", "name": "云平台EOS评估", "description": "评估云平台版本的EOS状态"}
  ],
  "evaluation_items": [
    {
      "dimension": "MT0",
      "name": "计划告警处理",
      "score": 52,
      "pass_rate": "80%",
      "level": "良",
      "result": "通过",
      "findings": "部分网元未及时处理告警，存在风险",
      "suggestions": "建议增加投资改造，增强动环监控"
    },
    {
      "dimension": "MT0",
      "name": "动环告警",
      "score": 55,
      "pass_rate": "82%",
      "level": "良",
      "result": "通过",
      "findings": "动环告警较少",
      "suggestions": "保持机房环境稳定"
    },
    {
      "dimension": "MT0",
      "name": "网元状态监测",
      "score": 58,
      "pass_rate": "89%",
      "level": "优",
      "result": "通过",
      "findings": "网元状态监测正常",
      "suggestions": "继续保持"
    }
  ],
  "overview": {
    "total_score": 72,
    "overall_pass_rate": "72%",
    "risk_level": "中",
    "summary": "网络架构高可用性评估得分72分，整体处于中等偏上水平。MT0、MT2表现良好，MT1、MT3、MT4需要重点关注优化。"
  }
}
```

## 字段说明

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `metadata` | object | 是 | 元数据信息 |
| `product_domains` | array | 是 | 产品域列表 |
| `network_elements` | array | 否 | 网元列表 |
| `evaluation_dimensions` | array | 是 | 评估维度定义 |
| `evaluation_items` | array | 是 | 评估项详情 |
| `overview` | object | 否 | 整体评估概览 |

## 评估维度说明

| 维度代码 | 维度名称 | 说明 |
|----------|----------|------|
| MT0 | 网元高稳评估 | 设备健康状态、告警数量、性能指标 |
| MT1 | 部署离散度评估 | 同城冗余部署、跨站点容灾 |
| MT2 | 组网架构高可用 | 主备配置、负载均衡、故障切换 |
| MT3 | 业务路由高可用 | 路由备份、切换时间、数据一致性 |
| MT4 | 网络容量均衡度 | 流量分布、资源利用率 |
| MT5 | 网元版本EOS | 版本生命周期、升级计划 |
| MT6 | 云平台EOS | OS版本、虚拟化平台版本 |
