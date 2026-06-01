# 多智能体服务

基于 LangGraph 的多智能体系统，提供总结、翻译、发邮件、制作 PPT、制作 Excel 功能。

## 功能

- **总结** - 智能总结任意内容为简洁、结构化的要点
- **翻译** - 中英文（或指定语言）高质量翻译
- **发邮件** - 自动整理收件人、主题和正文发送邮件
- **制作 PPT** - 根据内容自动生成演示文稿
- **制作 Excel** - 根据数据自动生成表格

## 架构

```
Supervisor (主管)
├── summarizer   - 总结专家
├── translator   - 翻译专家
├── emailer      - 邮件专家
├── ppt_maker    - PPT 专家
└── excel_maker  - Excel 专家
```

## 安装

```bash
pip install -r requirements.txt
```

## 配置

创建 `.env` 文件：

```env
# 火山方舟 API
ARK_API_KEY=你的密钥
MODEL_NAME=你的模型名
BASE_URL=https://ark.cn-beijing.volces.com/api/v3

# 邮件配置（可选，不发邮件可不配）
SMTP_HOST=smtp.gmail.com
SMTP_USER=你的邮箱
SMTP_PASS=你的密码
```

## 运行

```bash
uvicorn app:app --reload
```

访问：
- 前端页面：http://127.0.0.1:8000/static
- API 文档：http://127.0.0.1:8000/docs

## API 接口

| 接口 | 方法 | 说明 |
|------|------|------|
| `/chat` | POST | 普通对话（非流式） |
| `/chat/stream` | POST | 流式对话（SSE） |
| `/health` | GET | 健康检查 |
| `/api/files` | GET | 列出生成的文件 |
| `/files/{filename}` | GET | 下载生成的文件 |

## 示例

### Python 调用

```python
import requests

# 流式对话
response = requests.post(
    "http://localhost:8000/chat/stream",
    json={"message": "把下面这段话总结为三点：..."},
    stream=True
)
for line in response.iter_lines():
    if line.startswith("data: "):
        data = json.loads(line[6:])
        print(data.get("content", ""), end="")
```

### 请求示例

```json
// POST /chat/stream
{
  "message": "把这段话总结为三点：人工智能正在改变各行各业..."
}

// 响应 (SSE)
data: {"node": "supervisor", "content": "您好"}
data: {"node": "summarizer", "content": "已收到"}
data: {"node": "supervisor", "content": "为您总结如下："}
...
data: [DONE]
```

## 技术栈

- [LangGraph](https://langchain.ai/) - 多智能体框架
- [FastAPI](https://fastapi.tiangolo.com/) - Web 框架
- [python-pptx](https://python-pptx.readthedocs.io/) - PPT 生成
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel 生成