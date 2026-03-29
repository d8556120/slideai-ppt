"""
Stripe Payment Integration for AI PPT Generator SaaS.
Handles checkout sessions, webhooks, and subscription status.
"""

import os
from datetime import datetime

from fastapi import APIRouter, HTTPException, Request
from pydantic import BaseModel
from typing import Optional

router = APIRouter()

# ── Configuration ────────────────────────────────────────────────────────────

STRIPE_SECRET_KEY = os.environ.get("STRIPE_SECRET_KEY", "")
STRIPE_WEBHOOK_SECRET = os.environ.get("STRIPE_WEBHOOK_SECRET", "")
STRIPE_PRO_PRICE_ID = os.environ.get("STRIPE_PRO_PRICE_ID", "price_pro_placeholder")
STRIPE_TEAM_PRICE_ID = os.environ.get("STRIPE_TEAM_PRICE_ID", "price_team_placeholder")

DEMO_MODE = not STRIPE_SECRET_KEY

# Initialize Stripe if key is available
stripe = None
if not DEMO_MODE:
    import stripe as _stripe
    _stripe.api_key = STRIPE_SECRET_KEY
    stripe = _stripe


# ── In-Memory User Store (MVP) ───────────────────────────────────────────────

# Maps email -> { "plan": "free"|"pro"|"team", "stripe_customer_id": str, "subscription_id": str }
user_plans: dict = {}

# Maps email -> { "month": "YYYY-MM", "count": int }
usage_tracking: dict = {}


def get_user_plan(email: str) -> str:
    """Get the plan for a user. In demo mode, everyone gets 'pro'."""
    if DEMO_MODE:
        return "pro"
    return user_plans.get(email, {}).get("plan", "free")


def check_generation_allowed(email: str) -> tuple[bool, str]:
    """Check if a user is allowed to generate a presentation."""
    plan = get_user_plan(email)

    if plan in ("pro", "team"):
        return True, plan

    # Free plan: 3 per month
    current_month = datetime.now().strftime("%Y-%m")
    usage = usage_tracking.get(email, {"month": current_month, "count": 0})

    # Reset counter if new month
    if usage["month"] != current_month:
        usage = {"month": current_month, "count": 0}
        usage_tracking[email] = usage

    if usage["count"] >= 3:
        return False, "free_limit_reached"

    return True, plan


def record_generation(email: str):
    """Record a generation for rate limiting."""
    if not email:
        return
    current_month = datetime.now().strftime("%Y-%m")
    usage = usage_tracking.get(email, {"month": current_month, "count": 0})
    if usage["month"] != current_month:
        usage = {"month": current_month, "count": 0}
    usage["count"] += 1
    usage_tracking[email] = usage


# ── Request/Response Models ──────────────────────────────────────────────────

class CheckoutRequest(BaseModel):
    plan: str  # "pro" or "team"
    email: str
    success_url: Optional[str] = None
    cancel_url: Optional[str] = None


class CheckoutResponse(BaseModel):
    checkout_url: str
    session_id: str


class SubscriptionStatusResponse(BaseModel):
    email: str
    plan: str
    is_active: bool
    demo_mode: bool


# ── API Endpoints ────────────────────────────────────────────────────────────

@router.post("/api/checkout", response_model=CheckoutResponse)
async def create_checkout_session(req: CheckoutRequest):
    """Create a Stripe Checkout session for Pro or Team plan."""

    if req.plan not in ("pro", "team"):
        raise HTTPException(status_code=400, detail="Invalid plan. Choose 'pro' or 'team'.")

    if DEMO_MODE:
        # In demo mode, simulate a successful checkout
        return CheckoutResponse(
            checkout_url="/?success=true&demo=true",
            session_id="demo_session_" + req.email
        )

    price_id = STRIPE_PRO_PRICE_ID if req.plan == "pro" else STRIPE_TEAM_PRICE_ID

    # Default URLs based on the origin
    success_url = req.success_url or "/?success=true&session_id={CHECKOUT_SESSION_ID}"
    cancel_url = req.cancel_url or "/?canceled=true"

    try:
        session = stripe.checkout.Session.create(
            payment_method_types=["card"],
            mode="subscription",
            customer_email=req.email,
            line_items=[{
                "price": price_id,
                "quantity": 1,
            }],
            success_url=success_url,
            cancel_url=cancel_url,
            metadata={
                "plan": req.plan,
                "email": req.email,
            }
        )

        return CheckoutResponse(
            checkout_url=session.url,
            session_id=session.id
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to create checkout session: {str(e)}")


@router.post("/api/webhook")
async def stripe_webhook(request: Request):
    """Handle Stripe webhook events."""

    if DEMO_MODE:
        return {"status": "demo_mode", "message": "Webhooks disabled in demo mode"}

    payload = await request.body()
    sig_header = request.headers.get("stripe-signature", "")

    try:
        event = stripe.Webhook.construct_event(
            payload, sig_header, STRIPE_WEBHOOK_SECRET
        )
    except ValueError:
        raise HTTPException(status_code=400, detail="Invalid payload")
    except stripe.error.SignatureVerificationError:
        raise HTTPException(status_code=400, detail="Invalid signature")

    # Handle events
    if event["type"] == "checkout.session.completed":
        session = event["data"]["object"]
        email = session.get("customer_email") or session.get("metadata", {}).get("email", "")
        plan = session.get("metadata", {}).get("plan", "pro")
        customer_id = session.get("customer", "")
        subscription_id = session.get("subscription", "")

        if email:
            user_plans[email] = {
                "plan": plan,
                "stripe_customer_id": customer_id,
                "subscription_id": subscription_id,
            }
            print(f"[Stripe] User {email} subscribed to {plan} plan")

    elif event["type"] == "customer.subscription.deleted":
        subscription = event["data"]["object"]
        subscription_id = subscription.get("id", "")

        # Find and downgrade the user
        for email, data in user_plans.items():
            if data.get("subscription_id") == subscription_id:
                user_plans[email] = {
                    "plan": "free",
                    "stripe_customer_id": data.get("stripe_customer_id", ""),
                    "subscription_id": "",
                }
                print(f"[Stripe] User {email} subscription canceled, downgraded to free")
                break

    return {"status": "ok"}


@router.get("/api/subscription/status", response_model=SubscriptionStatusResponse)
async def subscription_status(email: str):
    """Check a user's subscription status."""

    if not email:
        raise HTTPException(status_code=400, detail="Email is required")

    plan = get_user_plan(email)

    return SubscriptionStatusResponse(
        email=email,
        plan=plan,
        is_active=plan in ("pro", "team"),
        demo_mode=DEMO_MODE,
    )
