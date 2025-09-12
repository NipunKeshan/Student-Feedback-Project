import os
import io
import sys
import csv
import zipfile
import platform
import tempfile
from pathlib import Path
from typing import List, Tuple

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import streamlit as st

# -----------------------------
# Page / Theme
# -----------------------------
st.set_page_config(page_title="Semester Feedback Dashboard", page_icon="📝", layout="wide")
st.markdown("""
<style>
/* Simple, clean styling */
.main > div { padding-top: 1rem; }
.block-container { padding-top: 1rem; }
h1, h2, h3 { font-family: system-ui, -apple-system, Segoe UI, Roboto, sans-serif; }
div[data-testid="stExpander"] div[role="button"] p {
  font-weight: 600 !important;
}
.plot-caption {
  font-size: 0.85rem; opacity: 0.8; margin-top: -0.25rem;
}
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
    # Count per ordered category
    counts = [(dq[dq[col] == label][col]).count() for label in LIKERT_ORDER]
    x = np.arange(len(LIKERT_ORDER))
    fig, ax = plt.subplots()
    ax.bar(LIKERT_ORDER, counts)
    ax.set_title(title)
    ax.set_xlabel("")  # keep clean
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

def zip_dir_bytes(dir_path: Path) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in dir_path.rglob("*"):
            if p.is_file():
                zf.write(p, arcname=p.relative_to(dir_path))
    buffer.seek(0)
    return buffer.read()

def open_email_with_attachments(subject: str, body: str, attachment_paths: List[Path]):
    system = platform.system().lower()

    if system.startswith("win"):
        try:
            import win32com.client as win32  # Requires: pip install pywin32
        except Exception as e:
            st.error("pywin32 is not installed. Run: `pip install pywin32`")
            return
        try:
            outlook = win32.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            mail.Subject = subject
            mail.Body = body
            for p in attachment_paths:
                mail.Attachments.Add(p.as_posix())
            mail.Display()  # opens compose window (does not send)
            st.success("Opened Outlook compose window.")
        except Exception as e:
            st.error(f"Unable to open Outlook compose window: {e}")
    elif system == "darwin":
        # macOS Apple Mail via AppleScript
        # We create a temporary AppleScript that opens a new message and attaches files
        try:
            # Escape file paths for AppleScript
            attachments_applescript = ""
            for p in attachment_paths:
                # Convert to POSIX path string and escape quotes
                attachments_applescript += f'set end of theAttachments to POSIX file "{p.as_posix()}"\n'

            script = f'''
            tell application "Mail"
                activate
                set theMessage to make new outgoing message with properties {{subject:"{subject}", content:"{body}\\n"}}
                tell theMessage
                    set visible to true
                    set theAttachments to {{}}
                    {attachments_applescript}
                    repeat with f in theAttachments
                        try
                            make new attachment with properties {{file name:f}} at after the last paragraph
                        end try
                    end repeat
                end tell
            end tell
            '''
            with tempfile.NamedTemporaryFile(suffix=".applescript", delete=False, mode="w") as tmp:
                tmp.write(script)
                tmp_path = tmp.name
            os.system(f'osascript "{tmp_path}"')
            st.success("Opened Apple Mail compose window.")
        except Exception as e:
            st.error(f"Unable to open Apple Mail compose window: {e}")
    else:
        st.info("Email compose automation is supported on Windows (Outlook) and macOS (Apple Mail) only.")

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

    with st.expander(f"📦 {feedback}  "
                     f"<span class='badge'>Lecturer: {lecturer}</span>  "
                     f"<span class='badge'>Module: {modulecode} - {modulename}</span>  "
                     f"<span class='badge'>Responses: {size}</span>", expanded=False):
        st.markdown("---")
        st.subheader("Likert Plots")

        # Grid layout for plots
        cols_per_row = 2
        for i, (col, label) in enumerate(zip(LIKERT_COLS, LIKERT_LABELS)):
            if i % cols_per_row == 0:
                row_cols = st.columns(cols_per_row, vertical_alignment="center")

            denom = size  # match original logic: percentage uses total SIZENO for the section
            fig, counts = plot_likert_bar(dq, col, label, denom_total=denom)

            # Save PNG to the per-combine folder
            filename = f"{i:02d}_{col}.png"
            png_path = save_fig(fig, out_dir, filename)

            # Show plot
            with row_cols[i % cols_per_row]:
                st.pyplot(fig)
                st.caption(f"<div class='plot-caption'>Saved: {filename}</div>", unsafe_allow_html=True)

        st.markdown("---")

        # Text comments (optional display + save to files like original)
        st.subheader("Comments")
        comments_tabs = st.tabs(TEXT_COMMENT_DESCRIPTIONS)
        for t_idx, (tcol, tdesc) in enumerate(zip(TEXT_COMMENT_COLUMNS, TEXT_COMMENT_DESCRIPTIONS)):
            # Gather non-empty, stripped unique comments
            comments_series = dq[tcol].dropna().astype(str).map(str.strip)
            comments = [c for c in comments_series.unique() if c]
            # Save to text file inside the same folder
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

        # Download ZIP (named after 'combine')
        zip_bytes = zip_dir_bytes(out_dir)
        zip_name = f"{safe_folder}.zip" if safe_folder else "plots.zip"
        st.download_button(
            label=f"📥 Download all plots & comments for '{feedback}'",
            data=zip_bytes,
            file_name=zip_name,
            mime="application/zip",
            help="Downloads a zip folder containing all PNG plots and the three comments text files."
        )

        # Email button (desktop only)
        st.markdown("#### Email these plots")
        st.caption("This opens your desktop email app with the PNGs attached (Windows: Outlook via MAPI, macOS: Apple Mail).")
        col_email1, col_email2 = st.columns([1, 3])
        with col_email1:
            do_email = st.button(f"✉️ Open compose window for '{feedback}'")
        with col_email2:
            default_subject = f"Lecturer feedback plots - {feedback}"
            subject = st.text_input("Email Subject", value=default_subject, key=f"subj_{safe_folder}")
            body = st.text_area("Email Body", value="Hi,\n\nPlease see the attached feedback plots.\n\nThanks.", key=f"body_{safe_folder}", height=100)

        if do_email:
            # Collect attachments (PNGs only, like your requirement)
            attachment_paths = sorted(out_dir.glob("*.png"))
            if not attachment_paths:
                st.warning("No PNG plots found to attach.")
            else:
                open_email_with_attachments(subject, body, attachment_paths)

# Footer
st.markdown("---")
st.caption("Built with Streamlit • Plots generated with Matplotlib • CSV format aligned with your original script")
