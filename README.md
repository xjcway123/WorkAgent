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
└── excel_maker - Excel 专家
```

## 项目结构

```
multiagent/
├── app.py              # FastAPI Web 服务
├── multi_agent.py      # 多智能体核心逻辑
├── frontend/
│   └── index.html      # 前端页面
├── skills/             # 技能模块
│   ├── graduation-defense-pptx/
│   └── excel_beautification/
├── output/             # 生成的文件目录
└── test.py             # 测试脚本
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

## 流式响应示例

```python
import requests
import json

response = requests.post(
    "http://localhost:8000/chat/stream",
    json={"message": "把这段话总结为三点：..."},
    stream=True
)
for line in response.iter_lines():
    if line.startswith("data: "):
        data = json.loads(line[6:])
        if data.get("content"):
            print(data["content"], end="")
        if data.get("node"):
            print(f"\n[当前智能体: {data['node']}]")
```

## 技术栈

- [LangGraph](https://langchain.ai/) - 多智能体框架
- [FastAPI](https://fastapi.tiangolo.com/) - Web 框架
- [python-pptx](https://python-pptx.readthedocs.io/) - PPT 生成
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel 生成