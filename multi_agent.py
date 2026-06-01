"""
LangGraph v1.0 多智能体系统 —— 起步骨架
功能:总结 / 翻译 / 发邮件 / 做 PPT / 做 Excel

架构:Supervisor(主管)路由到 5 个专家智能体。
依赖:pip install langchain langgraph langgraph-supervisor langchain-openai python-pptx openpyxl

环境变量(用 .env + python-dotenv 或直接 export):
    ARK_API_KEY   火山方舟密钥
    SMTP_HOST / SMTP_USER / SMTP_PASS   发邮件用(可选,不发邮件可不配)
"""

import os

from langchain_openai import ChatOpenAI
from langchain_core.tools import tool
from langchain.agents import create_agent              # ← v1.0 当前写法(取代 create_react_agent)
from langgraph_supervisor import create_supervisor
from pydantic import BaseModel, Field
from dotenv import load_dotenv
import smtplib
from email.mime.text import MIMEText
import re
from openpyxl import Workbook
from openpyxl.styles import Font

load_dotenv() 

# ----------------------------------------------------------------------
# 1. 模型:指向火山方舟(OpenAI 兼容),model 填你的接入点 ID
# ----------------------------------------------------------------------
model = ChatOpenAI(
    model=os.environ["MODEL_NAME"],                                # TODO: 换成你的接入点ID / 模型名
    base_url=os.environ["BASE_URL"],
    api_key=os.environ["ARK_API_KEY"],
    temperature=0,
)

OUTPUT_DIR = "./output"
os.makedirs(OUTPUT_DIR, exist_ok=True)



class EmailInput(BaseModel):
    to: str = Field(description="收件人邮箱地址")
    subject: str = Field(description="邮件主题")
    body: str = Field(description="邮件正文内容")


# ----------------------------------------------------------------------
# 2. 工具(tool):被对应的专家智能体调用。docstring 很重要,LLM 靠它判断怎么用
# ----------------------------------------------------------------------
@tool
def send_email(to: str, subject: str, body: str) -> str:
    """发送一封纯文本邮件。

    参数:
        to: 收件人邮箱地址,例如 someone@example.com
        subject: 邮件主题
        body: 邮件正文内容(纯文本)
    返回:
        发送结果说明(成功提示或失败原因)。
    """
    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = subject
    msg["From"] = os.environ["SMTP_USER"]
    msg["To"] = to

    port = int(os.getenv("SMTP_PORT", 465))   # 从 env 读端口,默认 465
    try:
        with smtplib.SMTP_SSL(os.environ["SMTP_HOST"], port) as s:
            s.login(os.environ["SMTP_USER"], os.environ["SMTP_PASS"])
            s.send_message(msg)
        return f"已成功发送邮件给 {to}(主题:{subject})"
    except smtplib.SMTPAuthenticationError:
        return "发送失败:邮箱认证失败,请检查 SMTP_USER 和授权码(SMTP_PASS)是否正确"
    except Exception as e:
        return f"发送失败:{e}"


@tool
def make_excel(filename: str, headers: list[str], rows: list[list]) -> str:
    """生成一个 Excel 文件并返回路径。

    参数:
        filename: 文件名(不含后缀),例如 "sales"
        headers: 表头列表,例如 ["月份", "销售额"]
        rows: 二维列表,每个子列表是一行,顺序对应 headers,
              例如 [["1月", 100], ["2月", 200]]
    返回:
        生成结果说明(成功路径或失败原因)。
    """
    # 文件名清洗:非法/路径字符替换为 _,防止报错或路径穿越
    safe_name = re.sub(r"[^\w\-]", "_", filename)
    path = os.path.join(OUTPUT_DIR, f"{safe_name}.xlsx")

    try:
        wb = Workbook()
        ws = wb.active

        # 写表头并加粗
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True)

        # 写数据,逐行校验列数与表头一致
        for i, row in enumerate(rows, start=1):
            if len(row) != len(headers):
                return f"生成失败:第 {i} 行有 {len(row)} 列,与表头 {len(headers)} 列不符"
            ws.append(row)

        wb.save(path)
        return f"Excel 已生成:{path}"
    except Exception as e:
        return f"生成 Excel 失败:{e}"


@tool
def make_ppt(filename: str, title: str, slides: list[dict]) -> str:
    """生成一个 PPT。title=封面标题, slides=内容页列表,每项形如 {"title": "页标题", "content": "页正文"}。返回文件路径。"""
    from pptx import Presentation

    prs = Presentation()
    # 封面页
    cover = prs.slides.add_slide(prs.slide_layouts[0])
    cover.shapes.title.text = title
    # 内容页
    for slide in slides:
        s = prs.slides.add_slide(prs.slide_layouts[1])
        s.shapes.title.text = slide.get("title", "")
        s.placeholders[1].text = slide.get("content", "")
    path = os.path.join(OUTPUT_DIR, f"{filename}.pptx")
    prs.save(path)
    return f"PPT 已生成:{path}"


# ----------------------------------------------------------------------
# 3. 专家智能体:总结/翻译是纯 LLM(无工具),其余各挂一个工具
# ----------------------------------------------------------------------
summarizer = create_agent(
    model,
    tools=[],
    system_prompt="你是总结专家。把用户提供的内容提炼成简洁、结构清晰的中文总结,抓住要点。",
    name="summarizer",
)

translator = create_agent(
    model,
    tools=[],
    system_prompt="你是翻译专家。根据用户要求在中英文(或指定语言)之间准确翻译,保持原意、语气和专有名词。",
    name="translator",
)

emailer = create_agent(
    model,
    tools=[send_email],
    system_prompt="你负责发送邮件。根据用户意图整理出收件人、主题和正文,然后调用 send_email 工具发送。",
    name="emailer",
)

ppt_maker = create_agent(
    model,
    tools=[make_ppt],
    system_prompt="你负责制作 PPT。把内容组织成一个封面标题 + 若干内容页(每页有标题和正文),再调用 make_ppt 生成。",
    name="ppt_maker",
)

excel_maker = create_agent(
    model,
    tools=[make_excel],
    system_prompt="你负责制作 Excel 表格。把数据整理成表头(headers)和数据行(rows),再调用 make_excel 生成。",
    name="excel_maker",
)


# ----------------------------------------------------------------------
# 4. 主管:负责把用户任务路由给合适的专家,支持多步串联
#    注意:supervisor 编排这一层对库版本较敏感,若组合报错见文件末尾说明
# ----------------------------------------------------------------------
supervisor = create_supervisor(
    agents=[summarizer, translator, emailer, ppt_maker, excel_maker],
    model=model,
    prompt=(
        "你是团队主管,负责把用户任务分派给合适的专家,自己不直接做事:\n"
        "- 内容总结/提炼 → summarizer\n"
        "- 语言翻译 → translator\n"
        "- 发送邮件 → emailer\n"
        "- 制作 PPT → ppt_maker\n"
        "- 制作 Excel → excel_maker\n"
        "若任务需要多步,按顺序依次调用。例如'把这段话总结后做成PPT':"
        "先让 summarizer 总结,再把结果交给 ppt_maker。\n"
        "所有专家完成后,向用户汇报最终结果。"
    ),
).compile()


# ----------------------------------------------------------------------
# 5. 运行示例
# ----------------------------------------------------------------------
def run(user_input: str):
    result = supervisor.invoke({"messages": [{"role": "user", "content": user_input}]})
    # 打印最终回复
    print(result["messages"][-1].content)
    return result


if __name__ == "__main__":
    # 示例 1:纯总结
    run("帮我把下面这段话总结成三点:人工智能正在改变各行各业……")

    # 示例 2:多步 —— 总结后做成 PPT
    # run("把这篇文章总结一下,然后做成一个3页的PPT:……")

    # 示例 3:做 Excel
    # run("帮我做一个销售表,列是 月份、销售额,数据:1月100,2月200,3月150")


# ======================================================================
# 排错说明
# ----------------------------------------------------------------------
# 1) 若 `from langchain.agents import create_agent` 报错:说明 langchain 版本偏旧,
#    升级:pip install -U langchain langgraph
# 2) 若 create_supervisor 与 create_agent 组合时报参数/类型错误(这层 API 版本敏感):
#    - 备选 A:把 create_agent 换成 langgraph.prebuilt.create_react_agent(会有弃用警告但能跑)
#    - 备选 B:不用 langgraph-supervisor 库,改用 StateGraph + 手写 handoff 工具自建主管
#    把你的报错和 `pip show langchain langgraph langgraph-supervisor` 版本发我,我帮你对症改。
# ======================================================================