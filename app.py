"""
用 FastAPI 封装多智能体系统(总结/翻译/发邮件/做PPT/做Excel)

运行:uvicorn app:app --reload
文档:http://127.0.0.1:8000/docs

依赖:在 multi_agent.py 的依赖之外,确保已装 fastapi、uvicorn:
    pip install fastapi "uvicorn[standard]"
"""

import json
import os

from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from multi_agent import supervisor   # 复用上一份文件里编译好的主管

app = FastAPI(title="多智能体服务", description="总结 / 翻译 / 发邮件 / 做PPT / 做Excel")

# 让生成的 PPT/Excel 可通过 /files/<文件名> 下载
os.makedirs("output", exist_ok=True)
app.mount("/files", StaticFiles(directory="output"), name="files")

# 前端页面放在 /static 路径下
app.mount("/static", StaticFiles(directory="frontend", html=True), name="frontend")


# ---------- 数据模型 ----------
class ChatRequest(BaseModel):
    message: str


class ChatResponse(BaseModel):
    reply: str


# ---------- 普通接口(一次性返回完整结果)----------
@app.post("/chat", response_model=ChatResponse)
async def chat(req: ChatRequest):
    try:
        result = await supervisor.ainvoke(
            {"messages": [{"role": "user", "content": req.message}]}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    return ChatResponse(reply=result["messages"][-1].content)


# ---------- 流式接口(逐 token 推送,SSE)----------
@app.post("/chat/stream")
async def chat_stream(req: ChatRequest):
    async def event_stream():
        try:
            async for token, meta in supervisor.astream(
                {"messages": [{"role": "user", "content": req.message}]},
                stream_mode="messages",
            ):
                if getattr(token, "content", ""):
                    data = {
                        "node": meta.get("langgraph_node"),  # 当前哪个智能体在输出
                        "content": token.content,
                    }
                    yield f"data: {json.dumps(data, ensure_ascii=False)}\n\n"
            yield "data: [DONE]\n\n"
        except Exception as e:
            yield f"data: {json.dumps({'error': str(e)}, ensure_ascii=False)}\n\n"

    return StreamingResponse(event_stream(), media_type="text/event-stream")


# ---------- 健康检查 ----------
@app.get("/health")
async def health():
    return {"status": "ok"}


# ---------- 文件列表 ----------
@app.get("/api/files")
async def list_files():
    """列出 output 目录下的 Excel 和 PPT 文件"""
    files = []
    if os.path.exists("output"):
        for f in os.listdir("output"):
            if f.endswith(('.xlsx', '.pptx')):
                files.append({
                    "name": f,
                    "url": f"/files/{f}",
                    "type": "xlsx" if f.endswith('.xlsx') else "pptx"
                })
    return {"files": files}