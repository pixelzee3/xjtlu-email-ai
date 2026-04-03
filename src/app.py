import asyncio
import sys

# Windows-specific event loop policy for Playwright compatibility
if sys.platform == "win32":
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

from pathlib import Path
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
import uvicorn
from typing import List, Optional
import os
import json

# Import functions from main.py without modifying it
from main import load_config, get_browser_page, search_emails, extract_full_body, call_llm

app = FastAPI()

# Global Exception Handler to ensure consistent JSON responses
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    return JSONResponse(
        status_code=500,
        content={"status": "error", "message": str(exc)},
    )

# Setup templates
templates = Jinja2Templates(directory="templates")

# Serve static files (i18n translations, etc.)
app.mount("/static", StaticFiles(directory="static"), name="static")

# Global state to hold browser instance and search results
class AppState:
    def __init__(self):
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        self.email_results = []  # Stores full email objects including locators
        self.config = {}

state = AppState()

# Pydantic models for request bodies
class SearchRequest(BaseModel):
    keyword: str

class ExecuteRequest(BaseModel):
    keyword: str
    instruction: str

@app.post("/api/execute")
async def execute_task(request: ExecuteRequest):
    await ensure_browser()
    
    keyword = request.keyword.strip()
    if not keyword:
        keyword = "basketball"
        
    instruction = request.instruction.strip()
    if not instruction:
        instruction = "请用自然语气总结这10封邮件，重点关注活动、课程作业和重要事项。"
    
    print(f"Executing task: search='{keyword}', instruction='{instruction}'")
    
    try:
        # 1. Search
        print(f"Searching for: {keyword}")
        emails = await search_emails(state.page, keyword)
        state.email_results = emails # Store in state just in case
        
        # 2. Extract bodies
        print(f"Extracting bodies for {len(emails)} emails...")
        full_bodies = []
        for i, email in enumerate(emails):
            print(f"Extracting body for email {i+1}...")
            body = await extract_full_body(state.page, email["locator"])
            full_bodies.append(f"=== 邮件 {i+1}: {email['subject']} ===\n日期: {email['date']}\n\n{body}\n\n{'='*80}\n")
            await asyncio.sleep(0.5)
            
        aggregated_text = "\n".join(full_bodies)
        
        # 3. Summarize
        from datetime import datetime
        summary_prompt = f"""
当前日期：{datetime.now().strftime('%Y-%m-%d')}
用户指令：{instruction}
请根据用户指令对以下邮件进行处理：

邮件内容：
{aggregated_text}
"""
        print("Calling LLM...")
        loop = asyncio.get_event_loop()
        summary = await loop.run_in_executor(None, call_llm, summary_prompt, state.config)
        
        # Prepare sanitized email list for frontend display
        sanitized_emails = []
        for i, email in enumerate(emails):
            sanitized_emails.append({
                "id": i,
                "subject": email.get("subject", ""),
                "date": email.get("date", ""),
                "href": email.get("href", "")
            })

        return {
            "status": "success", 
            "summary": summary,
            "emails": sanitized_emails
        }
        
    except Exception as e:
        error_msg = f"Task execution failed: {repr(e)}\n{traceback.format_exc()}"
        logger.error(error_msg)
        return JSONResponse(status_code=500, content={"status": "error", "message": str(e)})
@app.on_event("startup")
async def startup_event():
    # Load config on startup
    config_path = Path("config.json")
    if not config_path.exists():
        # Fallback if running from a different directory
        config_path = Path(__file__).parent / "config.json"
    
    state.config = load_config(config_path)
    print("Config loaded.")

@app.on_event("shutdown")
async def shutdown_event():
    # Cleanup browser resources
    if state.context:
        await state.context.close()
    if state.browser:
        await state.browser.close()
    if state.playwright:
        await state.playwright.stop()
    print("Browser resources released.")

import logging
import traceback

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

async def ensure_browser():
    """Ensure browser is running and page is ready."""
    if not state.page or state.page.is_closed():
        logger.info("Initializing browser...")
        try:
            state.playwright, state.browser, state.context, state.page = await get_browser_page(state.config)
        except Exception as e:
            error_msg = f"Failed to initialize browser: {repr(e)}\n{traceback.format_exc()}"
            logger.error(error_msg)
            raise HTTPException(status_code=500, detail=f"Failed to launch browser: {repr(e)}")

@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

# Remove old unused endpoints if they cause issues, but for now just fix the reference
# Actually I should remove the old endpoints since the frontend no longer uses them
# and SummarizeRequest was removed.

@app.post("/api/search")
async def search(request: SearchRequest):
    return {"status": "deprecated", "message": "Use /api/execute instead"}

@app.post("/api/summarize")
async def summarize(request: Request): # Changed to Request to avoid undefined model
    return {"status": "deprecated", "message": "Use /api/execute instead"}

if __name__ == "__main__":
    # Disable reload to avoid Windows subprocess loop issues
    # Use port 8001 to avoid conflicts with hung processes
    uvicorn.run("app:app", host="127.0.0.1", port=8001, reload=False, loop="asyncio")
