"""
AI PPT Generator SaaS - FastAPI Backend
Main application with API endpoints for generating, downloading, and managing presentations.
"""

import os
import time
import asyncio
from pathlib import Path
from contextlib import asynccontextmanager

# Load .env file if exists
env_path = Path(__file__).parent.parent / ".env"
if env_path.exists():
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            key, _, value = line.partition("=")
            os.environ.setdefault(key.strip(), value.strip())

from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from typing import Optional

from generator import generate_pptx, get_available_templates
from ai_content import generate_content
from payments import router as payments_router, check_generation_allowed, record_generation, DEMO_MODE


# ── Configuration ──────────────────────────────────────────────────────────────

GENERATED_DIR = Path(__file__).parent / "generated"
FRONTEND_DIR = Path(__file__).parent.parent / "frontend"
CLEANUP_INTERVAL = 600    # Check every 10 minutes
FILE_MAX_AGE = 3600       # 1 hour


# ── Cleanup background task ───────────────────────────────────────────────────

async def cleanup_old_files():
    """Periodically remove generated files older than FILE_MAX_AGE."""
    while True:
        await asyncio.sleep(CLEANUP_INTERVAL)
        try:
            now = time.time()
            if GENERATED_DIR.exists():
                for f in GENERATED_DIR.iterdir():
                    if f.suffix == ".pptx" and (now - f.stat().st_mtime) > FILE_MAX_AGE:
                        f.unlink()
                        print(f"Cleaned up: {f.name}")
        except Exception as e:
            print(f"Cleanup error: {e}")


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Manage application lifecycle - start cleanup task."""
    GENERATED_DIR.mkdir(parents=True, exist_ok=True)
    task = asyncio.create_task(cleanup_old_files())
    yield
    task.cancel()


# ── FastAPI App ────────────────────────────────────────────────────────────────

app = FastAPI(
    title="AI PPT Generator",
    description="Generate professional PowerPoint presentations with AI",
    version="1.0.0",
    lifespan=lifespan
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Mount Payments Router ─────────────────────────────────────────────────────

app.include_router(payments_router)


# ── Request/Response Models ────────────────────────────────────────────────────

class GenerateRequest(BaseModel):
    prompt: str = Field(..., min_length=5, max_length=2000, description="Describe the presentation you want")
    template: str = Field(default="business", description="Template theme: business, creative, or minimal")
    slide_count: int = Field(default=8, ge=5, le=15, description="Number of slides (5-15)")


class GenerateResponse(BaseModel):
    file_id: str
    download_url: str
    title: str
    subtitle: str
    slide_count: int
    slides_preview: list


# ── API Endpoints ──────────────────────────────────────────────────────────────

@app.post("/api/generate", response_model=GenerateResponse)
async def generate_presentation(req: GenerateRequest, request: Request):
    """Generate a PowerPoint presentation from a text prompt."""
    try:
        # Rate limiting based on user plan
        user_email = request.headers.get("X-User-Email", "")

        if user_email:
            allowed, status = check_generation_allowed(user_email)
            if not allowed:
                raise HTTPException(
                    status_code=429,
                    detail="Free plan limit reached (3 presentations/month). Upgrade to Pro for unlimited access."
                )

        # Generate content with AI (or demo mode)
        content = await generate_content(req.prompt, req.slide_count)

        # Generate the .pptx file
        file_id, filepath = generate_pptx(content, req.template)

        # Record usage for rate limiting
        if user_email:
            record_generation(user_email)

        # Build slides preview for frontend
        slides_preview = []
        for s in content.get("slides", []):
            slides_preview.append({
                "title": s.get("title", ""),
                "bullets": s.get("bullets", []),
            })

        return GenerateResponse(
            file_id=file_id,
            download_url=f"/api/download/{file_id}",
            title=content.get("title", "Presentation"),
            subtitle=content.get("subtitle", ""),
            slide_count=len(content.get("slides", [])) + 2,  # +title +closing
            slides_preview=slides_preview
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Generation failed: {str(e)}")


@app.get("/api/download/{file_id}")
async def download_presentation(file_id: str):
    """Download a generated .pptx file."""
    # Sanitize file_id to prevent path traversal
    if "/" in file_id or "\\" in file_id or ".." in file_id:
        raise HTTPException(status_code=400, detail="Invalid file ID")

    filepath = GENERATED_DIR / f"{file_id}.pptx"

    if not filepath.exists():
        raise HTTPException(status_code=404, detail="File not found or expired. Please generate again.")

    return FileResponse(
        path=str(filepath),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="presentation.pptx"
    )


@app.get("/api/templates")
async def list_templates():
    """List available presentation templates."""
    return {"templates": get_available_templates()}


@app.get("/api/health")
async def health_check():
    return {"status": "ok", "version": "2.0.0"}


# ── Serve Frontend ─────────────────────────────────────────────────────────────

@app.get("/")
async def serve_frontend():
    """Serve the main frontend page."""
    index_path = FRONTEND_DIR / "index.html"
    if index_path.exists():
        return HTMLResponse(content=index_path.read_text(encoding="utf-8"))
    return HTMLResponse(content="<h1>Frontend not found</h1><p>Place index.html in the frontend/ directory.</p>")


# Mount static files if directory exists
if FRONTEND_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(FRONTEND_DIR)), name="static")


# ── Run ────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import uvicorn
    print("\n" + "=" * 60)
    print("  AI PPT Generator SaaS")
    print("  http://localhost:8000")
    print("=" * 60)

    api_key = os.environ.get("OPENAI_API_KEY", "")
    if api_key:
        print("  Mode: AI-powered (OpenAI API connected)")
    else:
        print("  Mode: DEMO (set OPENAI_API_KEY for AI generation)")

    if DEMO_MODE:
        print("  Payments: DEMO (set STRIPE_SECRET_KEY for live payments)")
    else:
        print("  Payments: Stripe LIVE")

    print("=" * 60 + "\n")

    uvicorn.run(app, host="0.0.0.0", port=8000)
