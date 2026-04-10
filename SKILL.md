---
name: bid-content-generator
description: >-
  标书务虚方案自动生成。当用户需要为投标文件中的实施方案、项目管理方案、
  售后服务方案、培训方案、质量保障方案等"务虚"章节自动生成专业正文内容时，
  使用此skill。支持任何含分级标题和占位符的docx模板，通过调用OpenAI兼容的
  LLM API批量生成内容并正确回填，保持原始格式。触发关键词：标书、投标文件、
  务虚方案、实施方案、管理方案、售后服务、方案生成、标书填充、招标文件。
---

# 标书务虚方案自动生成

从招标文件框架docx中自动提取目录结构，调用LLM API为每个章节生成专业务虚内容，
正确回填到docx中，输出格式完整的投标方案文档。

## 工作原理

```
招标文件.docx（含框架目录+占位符）
    │
    ▼ extract_bid_structure.py
结构JSON（标题层级+占位符位置）
    │
    ▼ generate_content.py
内容JSON（每章节的正文内容）
    │
    ▼ fill_docx.py
完整方案.docx（格式正确+内容填充）
```

## 前置依赖

- **docx skill**：必须已安装（用于docx的解包/打包）
- **Python 3.8+**：运行脚本
- **lxml**：`pip install lxml`
- **LLM API Key**：需要有效的API密钥

## 使用流程

### 方式一：一键流水线（推荐）

直接调用 `scripts/run_pipeline.py`，三步合一：

```bash
python scripts/run_pipeline.py <招标文件.docx> <输出文件.docx> \
  --api-key sk-xxx \
  --model glm-4.5-air \
  --project-context "三亚崖州湾科技城车联网项目" \
  --industry "智慧交通"
```

参数说明：
- `--api-key`：LLM API密钥，也可通过环境变量 `LLM_API_KEY` 或文件 `~/.llm_api_key`
- `--model`：模型名称（默认 `glm-4.5-air`），支持任何OpenAI兼容模型
- `--api-base`：API地址（默认智谱AI），DeepSeek用 `https://api.deepseek.com/v1`
- `--project-context`：项目背景描述，帮助生成更贴合的内容
- `--industry`：行业领域
- `--chars-per-node`：每章节目标字数（默认700）
- `--batch-size`：每批处理数量（默认10）

### 方式二：分步执行（便于调试）

```bash
# 步骤1: 解析结构
python scripts/extract_bid_structure.py 招标文件.docx --output structure.json

# 步骤2: 生成内容（可分批次执行，支持断点续传）
python scripts/generate_content.py structure.json --model glm-4.5-air --batch 1
python scripts/generate_content.py structure.json --model glm-4.5-air --batch 2

# 步骤3: 回填到docx
python scripts/fill_docx.py 招标文件.docx generated_content.json 输出.docx
```

## Agent 操作指引

### 首次使用时的完整流程

1. **确认输入文件**：用户提供一个招标文件docx，要求文档中：
   - 包含分级标题（如"1 项目概述"、"1.1 建设目标"等）
   - 需要填充内容的章节标题后面有占位段落（如"【此处填写相关内容】"）
   - 标题使用Word内置Heading样式或包含编号模式

2. **确认API配置**：向用户获取以下信息（如果用户未提供，询问）：
   - API Key
   - 模型名称（推荐 glm-4.5-air，速率高、质量好）
   - API Base URL（不同厂商不同）
   - （可选）项目背景和行业领域

3. **执行流水线**：调用 `run_pipeline.py` 一键执行

4. **验证结果**：
   - 打开输出的docx，检查正文格式是否正确（宋体、首行缩进）
   - 检查是否有残留占位符
   - 检查内容连贯性和专业性

5. **如需修正**：
   - 如有残留占位符：检查占位符关键词是否匹配
   - 如格式不对：调整 `--body-font`、`--body-size` 等参数
   - 如内容质量不佳：换用更好的模型或增加 `--chars-per-node`

### 关键经验（务必遵守）

以下是从实际项目中总结的关键教训，违反这些规则会导致严重质量问题：

1. **不要硬往项目靠**：务虚方案要按通用知识体系写，不要把具体项目细节塞进每个章节
2. **正文不能有编号前缀**：生成的正文开头不要加"一、""1.""（1）"等编号
3. **公司名称用"我司"**：不要在生成内容中出现具体公司名称
4. **正文格式 ≠ 标题格式**：正文必须是普通段落样式，不能变成标题
5. **占位符替换策略**：找到标题后面紧跟的占位段落（无pPr），替换其文本并添加正文格式属性
6. **断点续传**：生成内容实时保存JSON，中途中断可以续传
7. **中文路径打包**：使用 `--validate false` 跳过GBK解码验证
8. **批量执行**：按10个一批调用API，避免限流

### 常见API配置

| 服务商 | API Base URL | 推荐模型 |
|--------|-------------|---------|
| 智谱AI | https://open.bigmodel.cn/api/paas/v4 | glm-4.5-air |
| DeepSeek | https://api.deepseek.com/v1 | deepseek-chat |
| 通义千问 | https://dashscope.aliyuncs.com/compatible-mode/v1 | qwen-plus |
| 硅基流动 | https://api.siliconflow.cn/v1 | Qwen/Qwen2.5-7B-Instruct |
| 本地Ollama | http://localhost:11434/v1 | 自定义 |

### 输入文档要求

最佳输入文档应满足：
- 使用Word内置Heading样式（heading1-5）或样式名包含"标题"
- 标题文本包含编号模式（如"1.1.1 xxx"）
- 需要填充内容的标题后面紧跟一个占位段落
- 文档整体结构完整（有封面、目录、各级标题）

如果输入文档的标题未使用标准样式，脚本会尝试通过编号模式自动推断层级。
