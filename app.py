import os
import io
import sys
import csv
import zipfile
import platform
import tempfile
from pathlib import Path
from typing import List, Tuple

import math
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import streamlit as st

# Email (Gmail SMTP)
import smtplib
import ssl
from email.message import EmailMessage

# -----------------------------
# Page / Theme
# -----------------------------
st.set_page_config(page_title="Semester Feedback Dashboard", page_icon="📝", layout="wide")
st.markdown("""
<style>
.main > div { padding-top: 1rem; }
.block-container { padding-top: 1rem; }
h1, h2, h3 { font-family: system-ui, -apple-system, Segoe UI, Roboto, sans-serif; }
div[data-testid="stExpander"] div[role="button"] p { font-weight: 600 !important; }
.plot-caption { font-size: 0.85rem; opacity: 0.8; margin-top: -0.25rem; }
hr { margin: 1rem 0; }
.badge { display: inline-block; padding: 4px 8px; border-radius: 999px; background: #EEF2FF; color: #3730A3; font-size: 12px; margin-left: 8px; }
</style>
""", unsafe_allow_html=True)

TITLE = "2025 Feb Semester Feedback Summary"

# -----------------------------
# Config (column names & labels)
# -----------------------------
TEXT_COMMENT_COLUMNS = ["commentslecturer", "commentsmodule", "commentsonline"]
TEXT_COMMENT_DESCRIPTIONS = ["Comments about Lecturer", "Comments about Module", "Comments about Online Delivery"]

LIKERT_COLS = [
    "explainations","organized","interesting","encourage","answers","consultation","additionalknowledge",
    "cameontime","easytounderstand","effective","teachingmaterial","tutoriallab","software","problemsolving",
    "facilities","videoclear","onlineforum","videostreaming","onlinenavigation","onlinesuccess"
]
LIKERT_LABELS = [
    "Explanations given by the lecturer were clear",
    "Conducted sessions in a well-planned and organized manner",
    "Made students interested and motivated",
    "Encouraged student participation and gave effective feedback",
    "Responded to students’ questions clearly and thoroughly",
    "Allocated sufficient time for student consultation",
    "Provided guidance about the subject and shared additional knowledge",
    "The lecturer came to class on time",
    "It was easy to understand and follow the lectures",
    "The lecturer was a very effective teacher",
    "Teaching materials and assignments were well-prepared",
    "Tutorial and lab activities were relevant to the lectures",
    "Software/hardware and other resources were available when needed",
    "Overall, the module enhanced problem-solving skills",
    "Teaching facilities provided for this module met my expectations",
    "Live/recorded videos were clear and audible",
    "An effective online forum was available",
    "Video streaming was high quality",
    "Easy to access/navigate the online teaching portal",
    "Online delivery successfully enabled me to continue with my studies"
]
LIKERT_ORDER = ["Strongly Agree", "Agree", "Neither", "Disagree", "Strongly Disagree"]

REQUIRED_COLUMNS = [
    "timestamp","refno","modulecode","modulename","lecturer","attendance","grade","explainations",
    "organized","interesting","encourage","answers","consultation","additionalknowledge","cameontime",
    "easytounderstand","effective","commentslecturer","teachingmaterial","tutoriallab","software",
    "problemsolving","facilities","commentsmodule","videoclear","onlineforum","videostreaming",
    "onlinenavigation","onlinesuccess","commentsonline","combine"
]

# Gmail nominal max ~25MB; base64 adds ~33%. Keep attachment <= ~17MB to be safe.
MAX_EMAIL_BYTES = 17 * 1024 * 1024  # 17 MiB

# -----------------------------
# Helpers
# -----------------------------
def validate_columns(df: pd.DataFrame) -> Tuple[bool, List[str]]:
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    return len(missing) == 0, missing

def annotate_percent(ax, counts, denom):
    for i, v in enumerate(counts):
        pct = (v * 100.0 / denom) if denom else 0.0
        ax.text(i, v + (max(counts) * 0.02 if counts else 0.1), f"{pct:.1f}%", ha="center", va="bottom", fontsize=9)

def plot_likert_bar(dq: pd.DataFrame, col: str, title: str, denom_total: int):
    counts = [(dq[dq[col] == label][col]).count() for label in LIKERT_ORDER]
    fig, ax = plt.subplots()
    ax.bar(LIKERT_ORDER, counts)
    ax.set_title(title)
    ax.set_xlabel("")
    ax.set_ylabel("Responses")
    annotate_percent(ax, counts, denom_total)
    ax.tick_params(axis='x', labelrotation=20)
    fig.tight_layout()
    return fig, counts

def save_fig(fig, out_dir: Path, filename: str) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / filename
    fig.savefig(out_path.as_posix(), bbox_inches="tight", pad_inches=0.3)
    plt.close(fig)
    return out_path

def zip_dir_bytes(dir_path: Path, only_ext: List[str] | None = None) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in dir_path.rglob("*"):
            if p.is_file():
                if only_ext:
                    if p.suffix.lower() not in [e.lower() for e in only_ext]:
                        continue
                zf.write(p, arcname=p.relative_to(dir_path))
    buffer.seek(0)
    return buffer.read()

def make_zip_from_files(files: List[Path], zip_path: Path) -> Path:
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, arcname=f.name)
    return zip_path

def base64_encoded_size(byte_len: int) -> int:
    """Return size after base64 (approx 4/3 growth)."""
    return math.ceil(byte_len / 3) * 4 + 50_000  # small header fudge

def batch_files_for_zip(files: List[Path], max_encoded_bytes: int) -> List[List[Path]]:
    """
    Greedy-pack files into batches so each resulting ZIP (after base64)
    stays under max_encoded_bytes. We approximate using sum of raw sizes
    and verify below when building actual zips.
    """
    # Conservative factor: ZIP compression is unknown; PNG is already compressed, TXT compresses well.
    # We'll estimate zip size roughly as sum(raw) * 1.05 (zip headers), then base64 growth.
    def est_encoded_zip_size(raw_sum: int) -> int:
        estimated_zip = int(raw_sum * 1.05)  # 5% overhead guess
        return base64_encoded_size(estimated_zip)

    batches = []
    cur, cur_raw = [], 0
    for f in files:
        sz = f.stat().st_size
        if est_encoded_zip_size(sz) > max_encoded_bytes:
            # Single file too big (unlikely for PNG/TXT). Put alone.
            if cur:
                batches.append(cur)
                cur, cur_raw = [], 0
            batches.append([f])
            continue
        if cur and est_encoded_zip_size(cur_raw + sz) > max_encoded_bytes:
            batches.append(cur)
            cur, cur_raw = [], 0
        cur.append(f)
        cur_raw += sz
    if cur:
        batches.append(cur)
    return batches

def refine_batches_to_fit(files: List[Path], max_encoded_bytes: int, safe_folder: str) -> List[Path]:
    """
    Build zips from approximate batches; if any zip still exceeds the budget, split further.
    Returns list of created zip file paths.
    """
    zips: List[Path] = []
    batches = batch_files_for_zip(files, max_encoded_bytes)
    tmp_root = Path(tempfile.gettempdir()) / f"email_zips_{safe_folder}"
    tmp_root.mkdir(parents=True, exist_ok=True)

    for idx, batch in enumerate(batches, start=1):
        # Try building a zip; if too big encoded, split batch and retry
        stack = [(idx, batch)]
        part_counter = 0
        while stack:
            tag, group = stack.pop()
            part_counter += 1
            zip_path = tmp_root / f"{safe_folder}_{tag}_{part_counter}.zip"
            make_zip_from_files(group, zip_path)
            raw = zip_path.stat().st_size
            enc = base64_encoded_size(raw)
            if enc <= max_encoded_bytes or len(group) == 1:
                zips.append(zip_path)
            else:
                # Split group roughly in half and retry
                mid = max(1, len(group) // 2)
                left, right = group[:mid], group[mid:]
                # push right then left so left is processed first
                stack.append((f"{tag}b", right))
                stack.append((f"{tag}a", left))
    return zips

def send_email_gmail_zips(to_emails: List[str], subject: str, body: str, zip_paths: List[Path]):
    sender = st.secrets["GMAIL_ADDRESS"]
    password = st.secrets["GMAIL_APP_PASSWORD"]

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender, password)
        total = len(zip_paths)
        for i, z in enumerate(zip_paths, start=1):
            msg = EmailMessage()
            suffix = f" (part {i}/{total})" if total > 1 else ""
            msg["From"] = sender
            msg["To"] = ", ".join(to_emails)
            msg["Subject"] = subject + suffix
            msg.set_content(body + (f"\n\nThis is {i} of {total}." if total > 1 else ""))
            with open(z, "rb") as fh:
                data = fh.read()
            msg.add_attachment(data, maintype="application", subtype="zip", filename=z.name)
            server.send_message(msg)

# -----------------------------
# UI: Header & Uploader
# -----------------------------
st.title("📝 Semester Feedback Dashboard")
st.caption("Upload the CSV (same column names as your original script). The app will group by **combine** and build plots & downloads per section.")

uploaded = st.file_uploader("Upload your feedback CSV", type=["csv"])

if not uploaded:
    st.info("Awaiting a CSV upload to begin...")
    st.stop()

# Read CSV
try:
    df = pd.read_csv(uploaded)
except Exception as e:
    st.error(f"Could not read CSV: {e}")
    st.stop()

valid, missing = validate_columns(df)
if not valid:
    st.error(f"Your CSV is missing required columns: {', '.join(missing)}")
    st.stop()

# Clean up 'combine' and prepare list
df['combine'] = df['combine'].astype(str).str.strip()
combine_list = df['combine'].dropna().unique()

# -----------------------------
# Summary header
# -----------------------------
st.markdown(f"### {TITLE}")
st.markdown(f"**Total Feedback Records:** {len(df)}")

# -----------------------------
# Process each 'combine'
# -----------------------------
for feedback in sorted(combine_list):
    dq = df.query("combine == @feedback").copy()
    if dq.empty:
        continue

    # Basic metadata for the section (first row)
    modulecode = str(dq["modulecode"].iloc[0]) if "modulecode" in dq else ""
    modulename = str(dq["modulename"].iloc[0]) if "modulename" in dq else ""
    lecturer = str(dq["lecturer"].iloc[0]) if "lecturer" in dq else ""
    size = dq.shape[0]

    # Create a stable temp folder for this combine's images
    safe_folder = "".join([c for c in feedback if c.isalnum() or c in ("-", "_", " ")])[:60].strip()
    base_dir = Path(tempfile.gettempdir()) / "feedback_plots"
    out_dir = base_dir / safe_folder
    out_dir.mkdir(parents=True, exist_ok=True)

    with st.expander(
        f"📦 {feedback}  "
        f"<span class='badge'>Lecturer: {lecturer}</span>  "
        f"<span class='badge'>Module: {modulecode} - {modulename}</span>  "
        f"<span class='badge'>Responses: {size}</span>", expanded=False
    ):
        st.markdown("---")
        st.subheader("Likert Plots")

        # Grid layout for plots
        cols_per_row = 2
        for i, (col, label) in enumerate(zip(LIKERT_COLS, LIKERT_LABELS)):
            if i % cols_per_row == 0:
                row_cols = st.columns(cols_per_row, vertical_alignment="center")

            denom = size  # percent uses total size for the section
            fig, counts = plot_likert_bar(dq, col, label, denom_total=denom)

            # Save PNG to the per-combine folder
            filename = f"{i:02d}_{col}.png"
            _ = save_fig(fig, out_dir, filename)

            # Show plot
            with row_cols[i % cols_per_row]:
                st.pyplot(fig)
                st.caption(f"<div class='plot-caption'>Saved: {filename}</div>", unsafe_allow_html=True)

        st.markdown("---")

        # Text comments
        st.subheader("Comments")
        comments_tabs = st.tabs(TEXT_COMMENT_DESCRIPTIONS)
        for t_idx, (tcol, tdesc) in enumerate(zip(TEXT_COMMENT_COLUMNS, TEXT_COMMENT_DESCRIPTIONS)):
            comments_series = dq[tcol].dropna().astype(str).map(str.strip)
            comments = [c for c in comments_series.unique() if c]
            # Save to text file
            txt_path = out_dir / f"{tcol}.txt"
            with open(txt_path, "w", encoding="utf-8", newline="") as f:
                for line in comments:
                    f.write(line + "\n")

            with comments_tabs[t_idx]:
                if comments:
                    for c in comments:
                        st.write(f"• {c}")
                else:
                    st.write("_No comments_")
                st.caption(f"<div class='plot-caption'>Saved: {tcol}.txt</div>", unsafe_allow_html=True)

        st.markdown("---")

        # Build a downloadable ZIP (PNG + TXT only)
        zip_bytes = zip_dir_bytes(out_dir, only_ext=[".png", ".txt"])
        zip_name = f"{safe_folder}.zip" if safe_folder else "plots.zip"
        st.download_button(
            label=f"📥 Download ZIP for '{feedback}'",
            data=zip_bytes,
            file_name=zip_name,
            mime="application/zip",
        )

        # -----------------------------
        # Instant send via Gmail (ZIP only)
        # -----------------------------
        st.markdown("#### Send ZIP via Gmail (instant)")
        st.caption("Attaches a ZIP (PNG plots + TXT comments). If too large, the app will split into multiple emails automatically.")

        col1, col2 = st.columns([2, 3])
        with col1:
            to_emails_raw = st.text_input("Recipient email(s) (comma-separated)", value="", key=f"to_{safe_folder}")
        with col2:
            subject = st.text_input("Email Subject", value=f"Lecturer feedback plots - {feedback}", key=f"subj_{safe_folder}")
        body = st.text_area(
            "Email Body",
            value="Hi,\n\nPlease see the attached feedback ZIP.\n\nThanks.",
            key=f"body_{safe_folder}",
            height=100
        )

        if st.button(f"🚀 Send ZIP now for '{feedback}'", key=f"send_{safe_folder}"):
            try:
                recipients = [x.strip() for x in to_emails_raw.split(",") if x.strip()]
                if not recipients:
                    st.warning("Please enter at least one recipient email.")
                else:
                    # Gather files to include in ZIPs
                    files_for_zip = sorted(list(out_dir.glob("*.png")) + list(out_dir.glob("*.txt")))
                    if not files_for_zip:
                        st.warning("No PNG/TXT files found to include in ZIP.")
                    else:
                        # Create size-safe zip parts
                        zip_paths = refine_batches_to_fit(files_for_zip, MAX_EMAIL_BYTES, safe_folder or "plots")
                        # Send one email per zip part
                        send_email_gmail_zips(recipients, subject, body, zip_paths)
                        st.success("Email(s) with ZIP sent successfully ✅")
            except KeyError:
                st.error("Gmail is not configured. Add GMAIL_ADDRESS and GMAIL_APP_PASSWORD to .streamlit/secrets.toml")
            except Exception as e:
                st.error(f"Failed to send email: {e}")

# Footer
st.markdown("---")
st.caption("Built with Streamlit • Plots generated with Matplotlib • Instant Gmail ZIP sending")
