import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.cluster import KMeans
from email.message import EmailMessage
import smtplib, os, io, tempfile, zipfile, hashlib

# =========================
# App / Secrets
# =========================
APP_URL = "https://excel-report-app-yzm5unptx7mgfebigwk8qi.streamlit.app"
BUSINESS_EMAIL = "j65146304@gmail.com"            # PayPal business email
EMAIL_ADDRESS  = st.secrets.get("EMAIL_ADDRESS", os.getenv("EMAIL_ADDRESS", ""))
EMAIL_PASSWORD = st.secrets.get("EMAIL_PASSWORD", os.getenv("EMAIL_PASSWORD", ""))

st.set_page_config(page_title="SheetGenius — AI Spreadsheet Automation", page_icon="📊", layout="centered")
st.title("📊 SheetGenius")
st.caption("AI-powered Excel/CSV cleaning, analysis, and charts — delivered to your inbox.")

# =========================
# Security settings & helpers
# =========================
MAX_SIZE_MB = 25
ALLOWED_EXTS = {".csv", ".xlsx", ".xls"}   # disallow .xlsm
MAX_ROWS = 1_000_000
MAX_COLS = 200
SUSPICIOUS_FORMULA_PREFIXES = ("=", "+", "-", "@")
SUSPICIOUS_STRINGS = (
    "WEBSERVICE(", "HYPERLINK(", "EXEC(", "SHELL(", "cmd.exe", "powershell",
    "URLDownloadToFile", "Auto_Open", "Workbook_Open"
)

def _ext(name: str) -> str:
    return ("." + name.rsplit(".", 1)[-1].lower()) if "." in name else ""

def human_size(n: int) -> str:
    return f"{n/1024/1024:.2f} MB"

def basic_magic_check_xlsx(first_bytes: bytes) -> bool:
    # XLSX is a ZIP: starts with PK
    return first_bytes[:2] == b"PK"

def hash_filelike(file) -> str:
    pos = file.tell()
    file.seek(0)
    h = hashlib.sha256()
    for chunk in iter(lambda: file.read(8192), b""):
        h.update(chunk)
    file.seek(pos)
    return h.hexdigest()

def is_safe_upload(uploaded_file) -> tuple[bool, str]:
    size_bytes = getattr(uploaded_file, "size", None)
    if size_bytes is not None and size_bytes > MAX_SIZE_MB * 1024 * 1024:
        return False, f"File too large ({human_size(size_bytes)}). Max is {MAX_SIZE_MB} MB."

    name = uploaded_file.name
    ext = _ext(name)
    if ext not in ALLOWED_EXTS:
        return False, f"Unsupported file type '{ext}'. Allowed: {', '.join(sorted(ALLOWED_EXTS))}."
    if ext == ".xlsm":
        return False, "Macro-enabled Excel (.xlsm) is not allowed."

    # quick sniff
    head = uploaded_file.read(4)
    uploaded_file.seek(0)
    if ext == ".xlsx" and not basic_magic_check_xlsx(head):
        return False, "XLSX did not look like a valid Excel (ZIP) file."

    file_hash = hash_filelike(uploaded_file)

    # sample read
    try:
        if ext == ".csv":
            df_sample = pd.read_csv(uploaded_file, nrows=500, on_bad_lines="skip", dtype=str, engine="python")
        else:
            df_sample = pd.read_excel(uploaded_file, nrows=500, dtype=str, engine="openpyxl")
    except Exception as e:
        uploaded_file.seek(0)
        return False, f"Could not parse sample of file: {e}"

    uploaded_file.seek(0)

    if df_sample.shape[1] > MAX_COLS:
        return False, f"Too many columns ({df_sample.shape[1]}). Max allowed: {MAX_COLS}."

    suspicious_hits = 0
    total_checked = 0
    for col in df_sample.columns:
        vals = df_sample[col].dropna().astype(str).head(500)
        total_checked += len(vals)
        for v in vals:
            v_strip = v.strip()
            if v_strip.startswith(SUSPICIOUS_FORMULA_PREFIXES):
                suspicious_hits += 1
            for needle in SUSPICIOUS_STRINGS:
                if needle.lower() in v_strip.lower():
                    suspicious_hits += 3
    if total_checked > 0 and (suspicious_hits / total_checked) > 0.1:
        return False, "File flagged for suspicious formula/content density."

    return True, f"OK (hash: {file_hash[:12]}…)"

def alert_admin(subject, body):
    # quiet failure (alerts should not crash app)
    try:
        if EMAIL_ADDRESS and EMAIL_PASSWORD:
            msg = EmailMessage()
            msg["Subject"] = subject
            msg["From"] = EMAIL_ADDRESS
            msg["To"] = EMAIL_ADDRESS
            msg.set_content(body)
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
                s.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
                s.send_message(msg)
    except Exception:
        pass

# =========================
# Payment paywall
# =========================
def render_paywall():
    st.markdown("### 💳 Step 1 — Choose a plan & pay with PayPal")
    buy_now_html = f"""
    <form action="https://www.paypal.com/cgi-bin/webscr" method="post" target="_top">
      <input type="hidden" name="cmd" value="_xclick">
      <input type="hidden" name="business" value="{BUSINESS_EMAIL}">
      <input type="hidden" name="item_name" value="SheetGenius Single Report">
      <input type="hidden" name="amount" value="25.00">
      <input type="hidden" name="currency_code" value="USD">
      <input type="hidden" name="return" value="{APP_URL}/?paid=1">
      <input type="hidden" name="cancel_return" value="{APP_URL}/?paid=0">
      <input type="submit" value="Pay $25 (One Report) via PayPal">
    </form>
    """
    subscribe_html = f"""
    <form action="https://www.paypal.com/cgi-bin/webscr" method="post" target="_top">
      <input type="hidden" name="cmd" value="_xclick-subscriptions">
      <input type="hidden" name="business" value="{BUSINESS_EMAIL}">
      <input type="hidden" name="item_name" value="SheetGenius Unlimited Plan">
      <input type="hidden" name="a3" value="100.00">
      <input type="hidden" name="p3" value="1">
      <input type="hidden" name="t3" value="M">
      <input type="hidden" name="src" value="1">
      <input type="hidden" name="sra" value="1">
      <input type="hidden" name="currency_code" value="USD">
      <input type="hidden" name="return" value="{APP_URL}/?paid=1">
      <input type="hidden" name="cancel_return" value="{APP_URL}/?paid=0">
      <input type="submit" value="Subscribe $100/month via PayPal">
    </form>
    """
    c1, c2 = st.columns(2)
    with c1: st.markdown(buy_now_html, unsafe_allow_html=True)
    with c2: st.markdown(subscribe_html, unsafe_allow_html=True)
    st.info("After successful payment, PayPal returns you here with ?paid=1 and the upload form unlocks.")

# =========================
# Processing helpers
# =========================
def process_df(df: pd.DataFrame, workdir: str) -> list[str]:
    outputs = []
    # summary
    summary_path = os.path.join(workdir, "summary.csv")
    df.describe(include="all").transpose().to_csv(summary_path)
    outputs.append(summary_path)

    # visual + clustering
    num = df.select_dtypes(include=[np.number]).columns.tolist()
    if len(num) >= 2:
        fig, ax = plt.subplots(figsize=(8,6))
        ax.scatter(df[num[0]], df[num[1]])
        ax.set_xlabel(num[0]); ax.set_ylabel(num[1]); ax.set_title("Scatterplot of Top Numeric Columns")
        plot_path = os.path.join(workdir, "scatterplot.png")
        fig.savefig(plot_path, bbox_inches="tight")
        plt.close(fig)
        outputs.append(plot_path)

        try:
            km = KMeans(n_clusters=3, n_init="auto", random_state=42)
            df2 = df.copy()
            df2["Cluster"] = km.fit_predict(df[num[:2]].fillna(0))
            cl_path = os.path.join(workdir, "clustered.csv")
            df2.to_csv(cl_path, index=False)
            outputs.append(cl_path)
        except Exception:
            pass

    return outputs

def send_email_with_attachments(to_email: str, files: list[str], subject: str, body: str, bcc_me: bool = True):
    if not EMAIL_ADDRESS or not EMAIL_PASSWORD:
        st.warning("Email secrets not set; skipping email. (You can still download the ZIP.)")
        return False
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = to_email
    if bcc_me:
        msg["Bcc"] = EMAIL_ADDRESS
    msg.set_content(body)
    for p in files:
        with open(p, "rb") as f:
            data = f.read()
        msg.add_attachment(data, maintype="application", subtype="octet-stream", filename=os.path.basename(p))
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            s.send_message(msg)
        return True
    except Exception as e:
        st.error(f"Email error: {e}")
        return False

# =========================
# Paywall gate
# =========================
paid_flag = st.query_params.get("paid", ["0"])[0] == "1"
if not paid_flag:
    render_paywall()
    st.stop()

st.success("✅ Payment confirmed. Upload is unlocked.")
st.markdown("### 📤 Step 2 — Upload your file & enter email")

uploaded = st.file_uploader("Upload CSV/XLSX (no .xlsm, up to 25MB)", type=["csv", "xlsx", "xls"])
client_email = st.text_input("Client email to receive the report")

if st.button("Submit & Process"):
    if not uploaded or not client_email:
        st.error("Please upload a file and enter an email.")
    else:
        # Security check BEFORE reading
        ok, reason = is_safe_upload(uploaded)
        if not ok:
            alert_admin("SheetGenius: blocked upload", f"Reason: {reason}\nFile: {uploaded.name}\nSize: {getattr(uploaded,'size',0)} bytes")
            st.error(f"Upload blocked: {reason}")
            st.stop()
        else:
            st.info(f"Security check passed: {reason}")

        with tempfile.TemporaryDirectory() as tmp:
            # Save to disk
            local_path = os.path.join(tmp, uploaded.name)
            with open(local_path, "wb") as f:
                f.write(uploaded.read())

            # Load full DF
            try:
                if uploaded.name.lower().endswith(".csv"):
                    df = pd.read_csv(local_path)
                else:
                    df = pd.read_excel(local_path)
            except Exception as e:
                st.error(f"Could not read the file: {e}")
                st.stop()

            # Hard caps
            if df.shape[0] > MAX_ROWS or df.shape[1] > MAX_COLS:
                st.error(f"File too large (rows: {df.shape[0]}, cols: {df.shape[1]}). Limits: {MAX_ROWS} rows, {MAX_COLS} cols.")
                st.stop()

            # Process
            outputs = process_df(df, tmp)

            # Zip outputs
            ts = pd.Timestamp.utcnow().strftime("%Y%m%d-%H%M%S")
            zip_path = os.path.join(tmp, f"sheetgenius_report_{ts}.zip")
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for p in outputs:
                    zf.write(p, arcname=os.path.basename(p))

            with open(zip_path, "rb") as zf:
                st.download_button("⬇️ Download results (ZIP)", zf, file_name=os.path.basename(zip_path), mime="application/zip")

            # Email
            mailed = send_email_with_attachments(
                to_email=client_email,
                files=[zip_path],
                subject="Your SheetGenius Report",
                body=("Thanks for using SheetGenius!\n\n"
                      "Attached: summary.csv, scatterplot.png (if available), clustered.csv (if applicable).\n"
                      "— SheetGenius"),
                bcc_me=True
            )
            if mailed:
                st.success(f"Report emailed to **{client_email}**.")
            else:
                st.warning("Email not sent (secrets missing or SMTP error). Use the download button above.")

<input type="hidden" name="return" value="https://YOUR-APP-URL/?paid=1">
<input type="hidden" name="cancel_return" value="https://YOUR-APP-URL/?paid=0">
