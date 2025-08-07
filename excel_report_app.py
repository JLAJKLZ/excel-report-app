import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.cluster import KMeans
import smtplib
from email.message import EmailMessage
import os
import tempfile

# Gmail credentials
EMAIL_ADDRESS = 'j65146304@gmail.com'
EMAIL_PASSWORD = 'dbij dhvh fwth yjgq'

# App UI
st.set_page_config(page_title="AI Spreadsheet Automation", layout="centered")
st.title("📊 AI-Powered Spreadsheet Automation")
st.write("Upload your Excel or CSV file and enter your email. You'll receive a cleaned data summary, visualizations, and optional clustering — all within 12–24 hours!")

# Upload file
uploaded_file = st.file_uploader("Upload your Excel or CSV file", type=["csv", "xlsx"])
email = st.text_input("Enter your email address")

# Submit button
if st.button("Submit and Process"):
    if uploaded_file is not None and email:
        with tempfile.TemporaryDirectory() as tmpdir:
            # Save uploaded file
            file_path = os.path.join(tmpdir, uploaded_file.name)
            with open(file_path, 'wb') as f:
                f.write(uploaded_file.read())

            # Load DataFrame
            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)

            # Generate summary
            summary_path = os.path.join(tmpdir, "summary.csv")
            df.describe(include='all').to_csv(summary_path)

            # Plot
            numeric_cols = df.select_dtypes(include='number').columns
            plot_path = os.path.join(tmpdir, "scatterplot.png")
            if len(numeric_cols) >= 2:
                plt.figure(figsize=(8, 6))
                sns.scatterplot(data=df, x=numeric_cols[0], y=numeric_cols[1])
                plt.title("Scatterplot of Top Numeric Columns")
                plt.savefig(plot_path)

            # Optional clustering
            cluster_path = os.path.join(tmpdir, "clustered.csv")
            if len(numeric_cols) >= 2:
                kmeans = KMeans(n_clusters=3)
                df['Cluster'] = kmeans.fit_predict(df[numeric_cols[:2]])
                df.to_csv(cluster_path, index=False)

            # Send email
            msg = EmailMessage()
            msg['Subject'] = 'Your AI-Processed Spreadsheet Report'
            msg['From'] = EMAIL_ADDRESS
            msg['To'] = email
            msg.set_content("Attached are your report summary, visualization, and clustering results.")

            for attachment in [summary_path, plot_path, cluster_path]:
                with open(attachment, 'rb') as f:
                    file_data = f.read()
                    file_name = os.path.basename(attachment)
                    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
                smtp.send_message(msg)

            st.success(f"✅ Report sent to {email} successfully!")
    else:
        st.error("Please upload a file and enter your email address.")
