"""
用 FastAPI 封装多智能体系统(总结/翻译/发邮件/做PPT/做Excel)

运行:uvicorn app:app --reload
文档:http://127.0.0.1:8000/docs

依赖:在 multi_agent.py 的依赖之外,确保已装 fastapi、uvicorn:
    pip install fastapi "uvicorn[standard]"
"""

import json
import os

from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from multi_agent import supervisor, astream_supervisor, _hitl_broker   # 复用上一份文件里编译好的主管

app = FastAPI(title="多智能体服务", description="总结 / 翻译 / 发邮件 / 做PPT / 做Excel")

# 让生成的 PPT/Excel 可通过 /files/<文件名> 下载
os.makedirs("output", exist_ok=True)
app.mount("/files", StaticFiles(directory="output"), name="files")

# 前端页面放在 /static 路径下
app.mount("/static", StaticFiles(directory="frontend", html=True), name="frontend")


# ---------- 数据模型 ----------
class ChatRequest(BaseModel):
    message: str


class RespondRequest(BaseModel):
    request_id: str
    response: str


# ---------- Human-in-the-Loop 响应端点 ----------
@app.post("/chat/respond")
async def chat_respond(req: RespondRequest):
    """前端收到 human_input_required 事件后,把用户选择 POST 回来以 resume agent 执行。"""
    ok = _hitl_broker.resolve(req.request_id, req.response)
    if not ok:
        # 可能是 request_id 已过期 / 已被 resolve / 客户端重复点击
        return {"ok": False, "reason": "no_pending_request"}
    return {"ok": True}


# ---------- 流式接口(逐事件推送,SSE)----------
@app.post("/chat")
async def chat(req: ChatRequest):
    """统一 SSE 端点:yield agent / token / tool_call / tool_result /
    human_input_required / done / error 事件。"""
    async def event_stream():
        try:
            async for evt in astream_supervisor(req.message):
                yield f"data: {json.dumps(evt, ensure_ascii=False)}\n\n"
        except Exception as e:
            yield f"data: {json.dumps({'event': 'error', 'message': str(e)}, ensure_ascii=False)}\n\n"

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