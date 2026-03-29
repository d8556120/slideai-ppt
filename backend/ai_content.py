"""
AI Content Generation Module
Generates structured PPT content from user prompts using OpenAI-compatible API.
Falls back to demo mode if no API key is configured.
"""

import os
import json
from typing import Optional

# Try importing openai - handle gracefully if not installed
try:
    from openai import OpenAI
except ImportError:
    OpenAI = None


def get_demo_content(prompt: str, slide_count: int = 8) -> dict:
    """Return demo content when no API key is configured."""
    base_slides = [
        {
            "title": "Market Overview",
            "bullets": [
                "Global market size reaching $4.2 trillion by 2027",
                "Year-over-year growth rate of 12.3%",
                "Key drivers: digital transformation, AI adoption",
                "Emerging markets contributing 40% of growth"
            ],
            "speaker_notes": "Start with the big picture to set context for the audience."
        },
        {
            "title": "Problem Statement",
            "bullets": [
                "Current solutions are fragmented and inefficient",
                "72% of businesses report productivity challenges",
                "Manual processes cost enterprises $1.8M annually",
                "Lack of integration between existing tools"
            ],
            "speaker_notes": "Emphasize the pain points that your audience can relate to."
        },
        {
            "title": "Our Solution",
            "bullets": [
                "AI-powered platform with end-to-end automation",
                "Seamless integration with 200+ enterprise tools",
                "Reduces operational costs by up to 60%",
                "Real-time analytics and actionable insights"
            ],
            "speaker_notes": "This is the core of your pitch - make it compelling."
        },
        {
            "title": "Key Features",
            "bullets": [
                "Intelligent workflow automation engine",
                "Natural language processing interface",
                "Advanced data visualization dashboard",
                "Enterprise-grade security and compliance"
            ],
            "speaker_notes": "Highlight the features that differentiate you from competitors."
        },
        {
            "title": "Technology Architecture",
            "bullets": [
                "Cloud-native microservices architecture",
                "Built on scalable Kubernetes infrastructure",
                "99.99% uptime SLA guarantee",
                "GDPR and SOC 2 Type II compliant"
            ],
            "speaker_notes": "Technical audience will appreciate the architectural decisions."
        },
        {
            "title": "Business Model",
            "bullets": [
                "SaaS subscription with tiered pricing",
                "Free tier for individual users (up to 3 projects)",
                "Professional plan: $49/month per user",
                "Enterprise: custom pricing with dedicated support"
            ],
            "speaker_notes": "Be transparent about pricing to build trust."
        },
        {
            "title": "Traction & Metrics",
            "bullets": [
                "50,000+ active users across 30 countries",
                "Monthly recurring revenue: $2.1M",
                "Net revenue retention rate: 135%",
                "Customer satisfaction score: 4.8/5.0"
            ],
            "speaker_notes": "Numbers speak louder than words - let the metrics tell the story."
        },
        {
            "title": "Competitive Landscape",
            "bullets": [
                "3x faster implementation than competitors",
                "Only solution offering full AI integration",
                "Lowest total cost of ownership in the market",
                "Awarded 'Best Innovation' at TechSummit 2025"
            ],
            "speaker_notes": "Focus on your unique advantages without disparaging competitors."
        },
        {
            "title": "Go-to-Market Strategy",
            "bullets": [
                "Direct sales team targeting Fortune 500",
                "Channel partnerships with major consultancies",
                "Content marketing and thought leadership",
                "Strategic partnerships with cloud providers"
            ],
            "speaker_notes": "Show that you have a clear plan to reach your target market."
        },
        {
            "title": "Roadmap & Vision",
            "bullets": [
                "Q2 2026: Launch enterprise marketplace",
                "Q3 2026: Expand to APAC region",
                "Q4 2026: AI copilot feature release",
                "2027: IPO preparation and Series C funding"
            ],
            "speaker_notes": "End with a forward-looking vision to inspire confidence."
        },
        {
            "title": "Team & Leadership",
            "bullets": [
                "Founded by ex-Google and ex-Microsoft leaders",
                "120+ team members across 5 global offices",
                "Advisory board includes Fortune 100 executives",
                "Deep expertise in AI, ML, and enterprise software"
            ],
            "speaker_notes": "People invest in teams - highlight your collective strength."
        },
        {
            "title": "Financial Projections",
            "bullets": [
                "Projected ARR: $50M by end of 2027",
                "Gross margin: 82% and improving",
                "Path to profitability by Q3 2027",
                "Capital efficient with 18-month runway"
            ],
            "speaker_notes": "Be realistic but optimistic with financial projections."
        }
    ]

    # Select slides based on requested count
    selected = base_slides[:min(slide_count - 2, len(base_slides))]  # Reserve 2 for title+closing

    topic = prompt[:80] if len(prompt) > 80 else prompt

    return {
        "title": f"Strategic Presentation: {topic}",
        "subtitle": "Comprehensive Analysis & Strategic Roadmap",
        "author": "AI PPT Generator",
        "slides": selected,
        "closing": {
            "title": "Thank You",
            "bullets": [
                "Questions & Discussion",
                "Contact: team@example.com",
                "Visit: www.example.com"
            ],
            "speaker_notes": "Open the floor for questions and provide contact information."
        }
    }


SYSTEM_PROMPT = """You are a professional presentation content generator. Given a topic or description, generate structured PowerPoint content in JSON format.

The JSON must have this exact structure:
{
  "title": "Main presentation title",
  "subtitle": "Subtitle or tagline",
  "author": "Presenter name or organization",
  "slides": [
    {
      "title": "Slide title",
      "bullets": ["Point 1", "Point 2", "Point 3", "Point 4"],
      "speaker_notes": "Brief notes for the presenter"
    }
  ],
  "closing": {
    "title": "Thank You / Closing title",
    "bullets": ["Contact info", "Next steps", "Call to action"],
    "speaker_notes": "Closing remarks"
  }
}

Rules:
- Each slide should have 3-5 bullet points
- Bullet points should be concise but informative (10-20 words each)
- Speaker notes should provide additional context (1-2 sentences)
- Content should be professional, well-organized, and flow logically
- Include relevant data points, statistics, or examples where appropriate
- Make the presentation tell a coherent story from start to finish
- Return ONLY valid JSON, no markdown or explanation"""


async def generate_content(prompt: str, slide_count: int = 8) -> dict:
    """
    Generate structured PPT content from a prompt.
    Uses OpenAI-compatible API if configured, otherwise returns demo content.
    """
    api_key = os.environ.get("OPENAI_API_KEY", "") or os.environ.get("ANTHROPIC_AUTH_TOKEN", "")
    base_url = os.environ.get("OPENAI_BASE_URL", "") or os.environ.get("ANTHROPIC_BASE_URL", "") or "https://api.openai.com/v1"

    # If no API key or openai not installed, use demo mode
    if not api_key or OpenAI is None:
        return get_demo_content(prompt, slide_count)

    try:
        client = OpenAI(api_key=api_key, base_url=base_url)

        user_message = f"""Create a professional presentation with exactly {slide_count - 2} content slides (plus title and closing slides) about the following topic:

{prompt}

Remember to return ONLY valid JSON matching the required structure."""

        response = client.chat.completions.create(
            model=os.environ.get("OPENAI_MODEL", "gpt-4o"),
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_message}
            ],
            temperature=0.7,
            max_tokens=4000
        )

        content = response.choices[0].message.content.strip()

        # Try to extract JSON from markdown code blocks if present
        if content.startswith("```"):
            lines = content.split("\n")
            json_lines = []
            in_block = False
            for line in lines:
                if line.startswith("```") and not in_block:
                    in_block = True
                    continue
                elif line.startswith("```") and in_block:
                    break
                elif in_block:
                    json_lines.append(line)
            content = "\n".join(json_lines)

        result = json.loads(content)
        return result

    except Exception as e:
        print(f"AI generation failed: {e}, falling back to demo mode")
        return get_demo_content(prompt, slide_count)
