# 暗盒笔记图片分析

这个项目会读取当前目录下的《暗盒笔记》xlsx，抽取“图片”列中的嵌入图片，并按原始 sheet、Excel 行号、序号、第几张图片保存。随后会调用 OpenAI 兼容的视觉 API，分析每张图片的光影、主旨和画面对象。

## 安装

```powershell
pip install -r requirements.txt
```

## 配置 API

```powershell
$env:OPENAI_API_KEY="你的 API Key"
```

也可以把 `.env.example` 复制成 `.env`，然后在 `.env` 里填写：

```text
OPENAI_API_KEY=你的 API Key
```

可选配置：

```powershell
$env:OPENAI_MODEL="gpt-4o-mini"
$env:OPENAI_BASE_URL="https://api.openai.com/v1"
```

如果你使用第三方 OpenAI 兼容服务，把 `OPENAI_BASE_URL` 和 `OPENAI_MODEL` 改成对方提供的值即可。

注意：本项目要分析图片，API 必须支持 `image_url` 视觉输入。DeepSeek 官方 API 当前不适合这个图片分析任务。

### 通义千问 Qwen 示例

```powershell
$env:OPENAI_API_KEY="你的百炼/DashScope API Key"
$env:OPENAI_BASE_URL="https://dashscope.aliyuncs.com/compatible-mode/v1"
$env:OPENAI_MODEL="qwen-vl-plus-latest"
python test.py --limit 3
```

### 豆包示例

豆包需要使用火山方舟/相关服务中的视觉理解模型或视觉端点，普通文本模型不能分析图片。

```powershell
$env:OPENAI_API_KEY="你的火山方舟 API Key"
$env:OPENAI_BASE_URL="https://ark.cn-beijing.volces.com/api/v3"
$env:OPENAI_MODEL="你的视觉模型名或推理接入点 ID"
python test.py --limit 3
```

## 先只抽取图片

```powershell
python test.py --extract-only
```

输出：

- `output/images/`：抽取出的图片
- `output/extracted_images_manifest.json`：图片与原始 sheet、序号、Excel 行号、第几张图片的匹配关系

## 调用 API 分析

```powershell
python test.py
```

输出：

- `output/analysis_results.jsonl`：逐张图片分析结果
- `output/analysis_results.xlsx`：便于查看的 Excel 汇总

测试前几张：

```powershell
python test.py --limit 3
```

如果遇到 `429 Too Many Requests`，说明 API 被限流或当前账号额度不足。可以降低速度：

```powershell
python test.py --sleep 10 --max-retries 8
```

脚本默认会读取已有的 `output/analysis_results.jsonl` 并跳过已成功的图片，适合中断后继续跑。
