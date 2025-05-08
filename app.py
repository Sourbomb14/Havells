import streamlit as st
import pandas as pd
import yagmail

# Set Streamlit page config
st.set_page_config(page_title="GSTIN Reconciliation App", layout="centered")

st.title("🧾 GSTIN Reconciliation & Notification Tool")

# Upload Excel files
company_file = st.file_uploader("Upload Company Excel File", type=["xlsx"])
payments_file = st.file_uploader("Upload Payments Excel File", type=["xlsx"])

# Email credentials
st.subheader("📧 Enter Gmail Credentials to Send Emails")
email_id = st.text_input("Gmail ID", placeholder="example@gmail.com")
email_password = st.text_input("App Password", type="password", placeholder="Enter your Gmail App Password")

# Button to trigger processing
if st.button("🚀 Reconcile and Show Results") and company_file and payments_file:
    # Read Excel files
    company_df = pd.read_excel(company_file, sheet_name=None, index_col="GSTIN of supplier")
    payments_df = pd.read_excel(payments_file, sheet_name=None, index_col="GSTIN of supplier")

    common_sheets = set(company_df.keys()).intersection(set(payments_df.keys()))

    for sheet in common_sheets:
        st.markdown(f"### 📄 Sheet: {sheet}")
        df_company = company_df[sheet]
        df_payments = payments_df[sheet]

        if 'Taxable Value (₹)' not in df_company.columns or 'Taxable Value (₹)' not in df_payments.columns:
            st.warning(f"❌ Missing 'Taxable Value (₹)' column in sheet: {sheet}")
            continue

        common_gstin = df_company.index.intersection(df_payments.index)

        taxable_comparison = pd.DataFrame({
            'Company_Taxable_Value': df_company.loc[common_gstin, 'Taxable Value (₹)'],
            'Payment_Taxable_Value': df_payments.loc[common_gstin, 'Taxable Value (₹)']
        })

        taxable_comparison['Difference'] = (
            taxable_comparison['Company_Taxable_Value'] - taxable_comparison['Payment_Taxable_Value']
        )

        st.success("📊 Difference in Taxable Value (₹) for common GSTINs")
        st.dataframe(taxable_comparison)

        unmatched_gstin = df_payments.index.difference(df_company.index)
        unmatched_records = df_payments.loc[unmatched_gstin]

        if not unmatched_records.empty:
            st.warning("📌 Records present in Payments sheet but missing in Company sheet")
            st.dataframe(unmatched_records)
            st.session_state["unmatched_records"] = unmatched_records
        else:
            st.info("✅ No unmatched records found.")

# Email sending logic
def send_email_to_supplier(email, name, gstin):
    try:
        subject = f"Missing Invoice Alert for GSTIN {gstin}"
        body = f"""
        Dear {name},

        Our system shows that your invoice with GSTIN {gstin} is present in our payment records
        but missing from our internal filing system.

        Kindly share the corresponding invoice/documents at the earliest so we can reconcile records.

        Regards,
        Finance Team
        """
        yag.send(to=email, subject=subject, contents=body)
        return f"✅ Email sent to {email}"
    except Exception as e:
        return f"❌ Failed to send email to {email}: {e}"

# Button to send emails
if st.button("📤 Send Emails to Unmatched Suppliers"):
    if not email_id or not email_password:
        st.error("❌ Please provide Gmail ID and App Password.")
    elif "unmatched_records" not in st.session_state:
        st.error("❌ Please upload and reconcile files first.")
    else:
        try:
            yag = yagmail.SMTP(user=email_id, password=email_password)
            unmatched_records = st.session_state["unmatched_records"]

            log_messages = []
            for gstin in unmatched_records.index:
                row = unmatched_records.loc[gstin]
                email = row.get('email')
                name = row.get('Trade/Legal name', 'Supplier')
                if pd.notnull(email) and isinstance(email, str):
                    log = send_email_to_supplier(email, name, gstin)
                    log_messages.append(log)
                else:
                    log_messages.append(f"⚠️ No valid email for GSTIN {gstin}")

            for msg in log_messages:
                st.write(msg)

        except Exception as e:
            st.error(f"❌ Gmail Authentication Failed: {e}")
