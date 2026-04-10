# Bid Content Generator Skill

标书务虚方案自动生成工具。从招标文件框架 docx 中自动提取目录结构，调用 LLM API 为每个章节生成专业务虚内容，正确回填到 docx 中，输出格式完整的投标方案文档。

## 特点

- 📄 支持任何含分级标题和占位符的 docx 模板
- 🤖 兼容所有 OpenAI 格式 API（智谱AI、DeepSeek、通义千问、Ollama 等）
- 🔄 断点续传：生成内容实时保存，中途中断可继续
- 📝 格式正确：正文宋体、首行缩进、1.5倍行距，符合标书排版规范
- ⚡ 批量执行：按 10 个一批调用 API，自动限流重试

## 工作流程

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

- **Python 3.8+**
- **lxml**：`pip install lxml`
- **docx skill**：必须已安装（用于 docx 的解包/打包）
- **LLM API Key**：需要有效的 API 密钥

## 快速使用

### 一键流水线（推荐）

```bash
python scripts/run_pipeline.py 招标文件.docx 输出方案.docx \
  --api-key sk-xxx \
  --model glm-4.5-air \
  --project-context "三亚崖州湾科技城车联网项目" \
  --industry "智慧交通"
```

### 分步执行（便于调试）

```bash
# 步骤1: 解析结构
python scripts/extract_bid_structure.py 招标文件.docx --output structure.json

# 步骤2: 生成内容（支持断点续传）
python scripts/generate_content.py structure.json --model glm-4.5-air --batch 1

# 步骤3: 回填到docx
python scripts/fill_docx.py 招标文件.docx generated_content.json 输出.docx
```

## 脚本说明

| 脚本 | 功能 | 关键参数 |
|------|------|----------|
| `extract_bid_structure.py` | 从 docx 提取目录结构 | `--placeholder`, `--min-level` |
| `generate_content.py` | 调用 LLM 批量生成内容 | `--api-key`, `--model`, `--chars-per-node`, `--batch` |
| `fill_docx.py` | 内容回填并打包 docx | `--body-font`, `--body-size`, `--line-spacing` |
| `run_pipeline.py` | 一键执行完整流水线 | 所有上述参数 + `--project-context`, `--industry` |

## 常见 API 配置

| 服务商 | API Base URL | 推荐模型 |
|--------|-------------|---------|
| 智谱AI | https://open.bigmodel.cn/api/paas/v4 | glm-4.5-air |
| DeepSeek | https://api.deepseek.com/v1 | deepseek-chat |
| 通义千问 | https://dashscope.aliyuncs.com/compatible-mode/v1 | qwen-plus |
| 硅基流动 | https://api.siliconflow.cn/v1 | Qwen/Qwen2.5-7B-Instruct |
| 本地Ollama | http://localhost:11434/v1 | 自定义 |

## 输入文档要求

- 使用 Word 内置 Heading 样式（heading1-5）或样式名包含"标题"
- 标题文本包含编号模式（如"1.1.1 xxx"）
- 需要填充内容的标题后面紧跟一个占位段落（如"【此处填写相关内容】"）
- 文档整体结构完整（有封面、目录、各级标题）

## 关键经验

1. **不要硬往项目靠**：务虚方案按通用知识体系写，不要把具体项目细节塞进每个章节
2. **正文不能有编号前缀**：生成的正文开头不要加"一、""1.""（1）"等编号
3. **公司名称用"我司"**：不要在生成内容中出现具体公司名称
4. **正文格式 ≠ 标题格式**：正文必须是普通段落样式，不能变成标题
5. **断点续传**：生成内容实时保存 JSON，中途中断可以续传
