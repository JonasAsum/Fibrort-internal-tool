"""The mail reader's agent identity in Margince: one Agent Seat Passport, kept in a local file.

Records the AI read out of the mails are written with it, so the CRM stamps them `agent:<passport id>` and
a deal closed won or lost by it waits in a human's approval inbox. Only the agent that staged an approval
can complete it afterwards (tested 2026-10-05: another passport gets "approval was staged by a different
agent"), so the passport is kept between runs rather than minted per run. It is bound to the human who
minted it, never has more rights than that human, expires after 90 days, and can be revoked in the CRM.
"""
import json
import os
from datetime import datetime, timedelta, timezone

from .client import ApiError, MargenceClient

STORE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".agent_passport.json")
LABEL = "mail reading (importer)"
TTL_HOURS = 2160  # the maximum, 90 days
RENEW_BEFORE = timedelta(days=7)


def _load(path: str) -> dict | None:
    try:
        with open(path) as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return None


def _usable(saved: dict | None, base_url: str, human_id: str) -> bool:
    if not saved or saved.get("base_url") != base_url or saved.get("on_behalf_of") != human_id:
        return False
    expires = datetime.fromisoformat(saved["expires_at"])
    return expires - datetime.now(timezone.utc) > RENEW_BEFORE


def agent_client(human: MargenceClient, store: str = STORE, cls=MargenceClient) -> tuple[MargenceClient, str]:
    """(client acting as the mail reader's agent, passport id). Mints a passport with the human session
    when there is none for this CRM and user, when it expires within a week, or when it was revoked."""
    human_id = human.get("/v1/me")["user"]["id"]
    saved = _load(store)
    if _usable(saved, human.base_url, human_id):
        agent = cls.as_agent(human.base_url, saved["token"])
        try:
            agent.get("/v1/approvals", params={"limit": 1})  # /v1/me is human-only; this answers any live agent
            return agent, saved["passport_id"]
        except ApiError as e:
            if e.response.status_code not in (401, 403):
                raise  # revoked or expired passports answer 401; anything else is a real problem
    minted = human.post("/v1/passports", json={"label": LABEL, "scopes": ["read", "write"], "ttl_hours": TTL_HOURS})
    saved = {"base_url": human.base_url, "passport_id": minted["passport_id"], "token": minted["token"],
             "on_behalf_of": minted["on_behalf_of"], "expires_at": minted["expires_at"]}
    fd = os.open(store, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)  # a credential: readable by this user only
    with os.fdopen(fd, "w") as fh:
        json.dump(saved, fh)
    return cls.as_agent(human.base_url, saved["token"]), saved["passport_id"]
