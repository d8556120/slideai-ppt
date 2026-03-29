"""
AI Content Generation Module
Generates structured PPT content from user prompts using OpenAI-compatible API.
Falls back to demo mode if no API key is configured.
"""

import os
import json
import random
import hashlib
from typing import Optional

# Try importing openai - handle gracefully if not installed
try:
    from openai import OpenAI
except ImportError:
    OpenAI = None


# ── Demo Content Sets ─────────────────────────────────────────────────────────

DEMO_CONTENT_SETS = {
    "business": [
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
    ],
    "education": [
        {
            "title": "Learning Landscape Today",
            "bullets": [
                "Global EdTech market projected to reach $404B by 2028",
                "Online learning adoption grew 300% since 2020",
                "Personalized learning improves outcomes by 30%",
                "Skills gap widening across STEM and digital literacy"
            ],
            "speaker_notes": "Set context with data on how education is rapidly transforming."
        },
        {
            "title": "Core Challenges in Education",
            "bullets": [
                "One-size-fits-all teaching fails diverse learners",
                "Teachers spend 50% of time on administrative tasks",
                "Student engagement drops 60% in passive lectures",
                "Assessment methods lag behind modern pedagogy"
            ],
            "speaker_notes": "Focus on challenges that resonate with educators and administrators."
        },
        {
            "title": "Our Approach",
            "bullets": [
                "Adaptive learning paths powered by AI algorithms",
                "Gamified modules that boost engagement by 45%",
                "Real-time progress tracking for students and teachers",
                "Multilingual content accessible on any device"
            ],
            "speaker_notes": "Present concrete solutions tied to each challenge mentioned."
        },
        {
            "title": "Curriculum Framework",
            "bullets": [
                "Aligned with national and international standards",
                "Project-based learning integrated into every module",
                "Cross-disciplinary STEM and humanities connections",
                "Regular content updates based on learner feedback"
            ],
            "speaker_notes": "Emphasize rigor and alignment with recognized standards."
        },
        {
            "title": "Technology Platform",
            "bullets": [
                "Cloud-based LMS with offline capability",
                "AI-driven content recommendation engine",
                "Interactive simulations and virtual labs",
                "Secure data handling compliant with FERPA and COPPA"
            ],
            "speaker_notes": "Address both innovation and data privacy concerns."
        },
        {
            "title": "Impact & Outcomes",
            "bullets": [
                "92% of pilot students improved test scores",
                "Teacher satisfaction rating: 4.7 out of 5.0",
                "Average session duration increased from 12 to 28 minutes",
                "Used in 500+ schools across 15 countries"
            ],
            "speaker_notes": "Let the outcomes data speak for itself."
        },
        {
            "title": "Partnerships & Ecosystem",
            "bullets": [
                "Strategic alliances with top universities",
                "Content partnerships with leading publishers",
                "Government grants in 8 countries for deployment",
                "Parent and community engagement portal"
            ],
            "speaker_notes": "Partnerships build credibility and extend reach."
        },
        {
            "title": "Scaling Plan",
            "bullets": [
                "Phase 1: Deepen penetration in existing markets",
                "Phase 2: Launch vocational and adult learning tracks",
                "Phase 3: Enter Southeast Asia and Africa",
                "Phase 4: AI tutor companion for 1-on-1 learning"
            ],
            "speaker_notes": "Show a clear, phased growth path."
        },
        {
            "title": "Investment & Financials",
            "bullets": [
                "Seed round of $5M fully deployed",
                "Seeking Series A of $20M for global expansion",
                "Current ARR: $3.2M with 15% month-over-month growth",
                "Unit economics positive across all cohorts"
            ],
            "speaker_notes": "Be transparent with numbers to build investor confidence."
        },
        {
            "title": "The Team",
            "bullets": [
                "Co-founders from Stanford and MIT education labs",
                "Former Google and Khan Academy engineers",
                "Advisors include UNICEF education directors",
                "Diverse team of 60 across engineering, content, and ops"
            ],
            "speaker_notes": "Highlight domain expertise and diversity."
        },
        {
            "title": "Vision for the Future",
            "bullets": [
                "Every learner deserves a personalized education path",
                "Technology as an equalizer, not a divider",
                "Goal: reach 10 million learners by 2028",
                "Building the infrastructure for lifelong learning"
            ],
            "speaker_notes": "End with an inspiring, mission-driven vision."
        },
        {
            "title": "Next Steps",
            "bullets": [
                "Pilot program available for new schools this quarter",
                "Free trial for individual educators",
                "Partnership inquiries welcome at partnerships@example.com",
                "Join our community of 10,000+ educators"
            ],
            "speaker_notes": "Give clear calls-to-action."
        }
    ],
    "technology": [
        {
            "title": "Industry Transformation",
            "bullets": [
                "AI and automation reshaping every sector",
                "Cloud spending surpassed $600B globally in 2025",
                "Cybersecurity threats increasing 38% year over year",
                "Developer productivity tools market growing at 25% CAGR"
            ],
            "speaker_notes": "Establish the macro trends driving technology adoption."
        },
        {
            "title": "The Problem We Solve",
            "bullets": [
                "Engineering teams waste 40% of time on repetitive tasks",
                "Legacy systems create $2.5 trillion in technical debt globally",
                "Deployment failures cost enterprises $400K per incident",
                "Cross-team collaboration remains a bottleneck"
            ],
            "speaker_notes": "Quantify pain points with concrete data."
        },
        {
            "title": "Product Overview",
            "bullets": [
                "Unified developer platform with AI-assisted coding",
                "Automated CI/CD pipelines with zero-config setup",
                "Intelligent code review and security scanning",
                "Real-time collaboration across distributed teams"
            ],
            "speaker_notes": "Position the product as a comprehensive developer solution."
        },
        {
            "title": "Architecture & Stack",
            "bullets": [
                "Rust-based core engine for maximum performance",
                "WebAssembly runtime for cross-platform plugins",
                "Event-driven architecture handling 100K requests/second",
                "End-to-end encryption with zero-knowledge design"
            ],
            "speaker_notes": "Technical depth builds credibility with engineering audiences."
        },
        {
            "title": "Developer Experience",
            "bullets": [
                "Setup time reduced from days to under 5 minutes",
                "IDE extensions for VS Code, JetBrains, and Neovim",
                "CLI-first design with comprehensive API coverage",
                "Documentation scored 4.9/5 by developer community"
            ],
            "speaker_notes": "Developer experience is the key differentiator."
        },
        {
            "title": "Security & Compliance",
            "bullets": [
                "SOC 2 Type II and ISO 27001 certified",
                "Automated vulnerability scanning on every commit",
                "Role-based access control with SSO integration",
                "Audit logs and compliance reporting built-in"
            ],
            "speaker_notes": "Security is non-negotiable for enterprise customers."
        },
        {
            "title": "Performance Benchmarks",
            "bullets": [
                "Build times reduced by 70% compared to alternatives",
                "99.995% platform availability over the past 12 months",
                "Median API response time: 23 milliseconds",
                "Supports repositories with 10M+ lines of code"
            ],
            "speaker_notes": "Hard numbers differentiate you from competitors making vague claims."
        },
        {
            "title": "Customer Success Stories",
            "bullets": [
                "Fortune 100 bank reduced deployment time by 85%",
                "Healthcare startup achieved HIPAA compliance in 2 weeks",
                "E-commerce giant handles 1M+ daily deploys",
                "Open-source project gained 50K stars using our platform"
            ],
            "speaker_notes": "Social proof from recognizable names builds trust."
        },
        {
            "title": "Growth Metrics",
            "bullets": [
                "200,000+ developers on the platform",
                "ARR: $15M with 120% net dollar retention",
                "Enterprise customers: 85 and growing",
                "Community contributions: 5,000+ plugins"
            ],
            "speaker_notes": "Growth trajectory matters more than absolute numbers at this stage."
        },
        {
            "title": "Pricing & Plans",
            "bullets": [
                "Free tier: unlimited public projects",
                "Pro: $19/dev/month with advanced features",
                "Team: $39/dev/month with admin controls",
                "Enterprise: custom pricing with SLA and support"
            ],
            "speaker_notes": "Transparent pricing reduces friction in the sales process."
        },
        {
            "title": "Product Roadmap",
            "bullets": [
                "Q2 2026: AI pair programming assistant launch",
                "Q3 2026: On-premise deployment option",
                "Q4 2026: Advanced analytics and DORA metrics",
                "2027: Platform SDK for custom integrations"
            ],
            "speaker_notes": "Show you have a clear vision and execution plan."
        },
        {
            "title": "Join Our Ecosystem",
            "bullets": [
                "Start free at platform.example.com",
                "Developer community on Discord: 25K+ members",
                "Weekly office hours and monthly webinars",
                "Contact sales@example.com for enterprise inquiries"
            ],
            "speaker_notes": "End with clear next steps and entry points."
        }
    ]
}


def _select_demo_set(prompt: str) -> str:
    """Select a demo content set based on prompt keywords."""
    prompt_lower = prompt.lower()

    education_keywords = ["education", "learning", "school", "university", "student",
                          "teacher", "course", "curriculum", "training", "academic"]
    technology_keywords = ["technology", "software", "developer", "code", "platform",
                           "engineering", "devops", "cloud", "api", "infrastructure",
                           "saas", "app", "security", "cyber"]

    edu_score = sum(1 for kw in education_keywords if kw in prompt_lower)
    tech_score = sum(1 for kw in technology_keywords if kw in prompt_lower)

    if edu_score > tech_score and edu_score > 0:
        return "education"
    elif tech_score > edu_score and tech_score > 0:
        return "technology"

    # Use a hash of the prompt to pick deterministically but varied
    hash_val = int(hashlib.md5(prompt.encode()).hexdigest(), 16)
    keys = list(DEMO_CONTENT_SETS.keys())
    return keys[hash_val % len(keys)]


def get_demo_content(prompt: str, slide_count: int = 8, language: str = "English",
                     tone: str = "professional", audience: str = "general",
                     include_statistics: bool = True) -> dict:
    """Return demo content when no API key is configured."""
    content_key = _select_demo_set(prompt)
    base_slides = DEMO_CONTENT_SETS[content_key]

    # Select slides based on requested count
    selected = base_slides[:min(slide_count - 2, len(base_slides))]  # Reserve 2 for title+closing

    topic = prompt[:80] if len(prompt) > 80 else prompt

    # Adjust title phrasing based on tone
    tone_prefixes = {
        "professional": "Strategic Presentation",
        "casual": "Quick Look",
        "academic": "Research Overview",
        "creative": "Big Ideas",
    }
    prefix = tone_prefixes.get(tone, "Strategic Presentation")

    return {
        "title": f"{prefix}: {topic}",
        "subtitle": "Comprehensive Analysis & Strategic Roadmap",
        "author": "SlideAI Generator",
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


def _build_system_prompt(language: str = "English", tone: str = "professional",
                         audience: str = "general", include_statistics: bool = True) -> str:
    """Build a detailed system prompt incorporating all parameters."""

    tone_guidance = {
        "professional": "Use a formal, polished, and business-appropriate tone. Write clearly and assertively.",
        "casual": "Use a friendly, conversational, and approachable tone. Keep it light but informative.",
        "academic": "Use a scholarly, evidence-based, and precise tone. Reference methodology and cite data where possible.",
        "creative": "Use an imaginative, bold, and engaging tone. Employ vivid language and storytelling techniques.",
    }

    audience_guidance = {
        "investors": "Tailor content for investors and stakeholders. Emphasize ROI, market opportunity, traction metrics, competitive moats, and financial projections.",
        "team": "Tailor content for internal team members. Focus on actionable plans, responsibilities, timelines, and collaboration points.",
        "customers": "Tailor content for customers and prospects. Highlight benefits, use cases, testimonials, and value propositions.",
        "general": "Tailor content for a general audience. Balance depth and accessibility, define jargon, and use relatable examples.",
    }

    stats_instruction = ""
    if include_statistics:
        stats_instruction = "\n- Include concrete statistics, percentages, dollar figures, or data points in bullet points wherever relevant. Use realistic and plausible numbers."
    else:
        stats_instruction = "\n- Focus on qualitative insights and strategic points rather than specific numbers or statistics."

    return f"""You are SlideAI, an expert presentation content generator. Given a topic or description, generate structured PowerPoint content in JSON format.

LANGUAGE: Generate ALL content (titles, bullets, speaker notes) in {language}.

TONE: {tone_guidance.get(tone, tone_guidance["professional"])}

AUDIENCE: {audience_guidance.get(audience, audience_guidance["general"])}

The JSON must have this exact structure:
{{
  "title": "Main presentation title",
  "subtitle": "Subtitle or tagline",
  "author": "Presenter name or organization",
  "slides": [
    {{
      "title": "Slide title",
      "bullets": ["Point 1", "Point 2", "Point 3", "Point 4"],
      "speaker_notes": "Brief notes for the presenter"
    }}
  ],
  "closing": {{
    "title": "Thank You / Closing title",
    "bullets": ["Contact info", "Next steps", "Call to action"],
    "speaker_notes": "Closing remarks"
  }}
}}

Rules:
- Each slide should have 3-5 bullet points
- Bullet points should be concise but informative (10-20 words each)
- Speaker notes should provide additional context (1-2 sentences)
- Content should be {tone}, well-organized, and flow logically{stats_instruction}
- Make the presentation tell a coherent story from start to finish
- Ensure the content is appropriate for the target audience: {audience}
- Return ONLY valid JSON, no markdown or explanation"""


async def generate_content(prompt: str, slide_count: int = 8,
                           language: str = "English", tone: str = "professional",
                           audience: str = "general",
                           include_statistics: bool = True) -> dict:
    """
    Generate structured PPT content from a prompt.
    Uses OpenAI-compatible API if configured, otherwise returns demo content.

    Args:
        prompt: The topic or description for the presentation
        slide_count: Number of slides to generate (5-15)
        language: Language for the content (default: English)
        tone: Tone of the content (professional/casual/academic/creative)
        audience: Target audience (investors/team/customers/general)
        include_statistics: Whether to include statistics in the content

    Returns:
        dict with structured presentation content
    """
    api_key = os.environ.get("OPENAI_API_KEY", "") or os.environ.get("ANTHROPIC_AUTH_TOKEN", "")
    base_url = os.environ.get("OPENAI_BASE_URL", "") or os.environ.get("ANTHROPIC_BASE_URL", "") or "https://api.openai.com/v1"

    # If no API key or openai not installed, use demo mode
    if not api_key or OpenAI is None:
        return get_demo_content(prompt, slide_count, language, tone, audience, include_statistics)

    system_prompt = _build_system_prompt(language, tone, audience, include_statistics)

    user_message = f"""Create a {tone} presentation with exactly {slide_count - 2} content slides (plus title and closing slides) about the following topic:

{prompt}

Target audience: {audience}
Language: {language}
{"Include concrete statistics and data points." if include_statistics else "Focus on qualitative insights."}

Remember to return ONLY valid JSON matching the required structure."""

    max_attempts = 2  # Retry once on failure

    for attempt in range(max_attempts):
        try:
            client = OpenAI(api_key=api_key, base_url=base_url)

            response = client.chat.completions.create(
                model=os.environ.get("OPENAI_MODEL", "gpt-4o"),
                messages=[
                    {"role": "system", "content": system_prompt},
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
            print(f"AI generation attempt {attempt + 1}/{max_attempts} failed: {e}")
            if attempt < max_attempts - 1:
                print("Retrying...")
                continue
            else:
                print("All attempts failed, falling back to demo mode")
                return get_demo_content(prompt, slide_count, language, tone, audience, include_statistics)
