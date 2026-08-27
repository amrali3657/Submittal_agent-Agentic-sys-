"""
notify.py

Builds and sends the team notification email: new items, status changes, a
current status breakdown, and links to any engineer-comment files pulled
down this run. Same Gmail SMTP + app-password pattern as the Training
Matrix Monitor, kept consistent on purpose.
"""
import smtplib
from collections import Counter
from email.mime.text import MIMEText


def _status_breakdown(all_submittals: list[dict]) -> list[str]:
    counts = Counter(s["status"] for s in all_submittals if s.get("status"))
    if not counts:
        return []
    lines = ["Current status breakdown:"]
    for status, n in sorted(counts.items(), key=lambda kv: -kv[1]):
        lines.append(f"  - {status}: {n}")
    return lines


def notify_changes(
    cfg,
    added: list[str],
    changed: list[str],
    all_submittals: list[dict],
    uploaded_responses: dict[str, list[str]] | None = None,
    dropbox_shared_link_base: str = "",
) -> None:
    if not (cfg.GMAIL_ADDRESS and cfg.GMAIL_APP_PASSWORD and cfg.NOTIFY_TO):
        return
    if not added and not changed and not uploaded_responses:
        return

    lines = ["Submittal log sync ran — here's the summary:", ""]
    if added:
        lines.append(f"New submittals ({len(added)}): {', '.join(added)}")
    if changed:
        lines.append(f"Status changes ({len(changed)}): {', '.join(changed)}")
    lines.append("")
    lines.extend(_status_breakdown(all_submittals))

    if uploaded_responses:
        lines.append("")
        lines.append("New Jacobs/designer response files pulled from SharePoint:")
        for tul_no, paths in uploaded_responses.items():
            for p in paths:
                link = f"{dropbox_shared_link_base}{p}" if dropbox_shared_link_base else p
                lines.append(f"  - [{tul_no}] {link}")

    body = "\n".join(lines)

    msg = MIMEText(body)
    msg["Subject"] = f"Submittal Log Update — {len(added)} new, {len(changed)} changed"
    msg["From"] = cfg.GMAIL_ADDRESS
    msg["To"] = cfg.NOTIFY_TO

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(cfg.GMAIL_ADDRESS, cfg.GMAIL_APP_PASSWORD)
        server.sendmail(cfg.GMAIL_ADDRESS, [cfg.NOTIFY_TO], msg.as_string())
