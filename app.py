import streamlit as st
import pandas as pd
import yagmail

st.set_page_config(page_title="GSTIN Reconciliation Tool", layout="wide")
st.title("📊 GSTIN Reconciliation and Email Notifier")

st.markdown("Upload **Company** and **Payments** Excel files for comparison.")

# Upload files
company_file = st.file_uploader("Upload Company Excel File", type=["xlsx"])
payments_file = st.file_uploader("Upload Payments Excel File", type=["xlsx"])

if company_file and payments_file:
    company_data = pd.read_excel(company_file, sheet_name=None, index_col='GSTIN of supplier')
    payments_data = pd.read_excel(payments_file, sheet_name=None, index_col='GSTIN of supplier')

    common_sheets = set(company_data.keys()).intersection(set(payments_data.keys()))
    
    for sheet in common_sheets:
        st.header(f"🗂️ Sheet: {sheet}")

        df_company = company_data[sheet]
        df_payments = payments_data[sheet]

        if 'Taxable Value (₹)' not in df_company.columns or 'Taxable Value (₹)' not in df_payments.columns:
            st.warning("Missing 'Taxable Value (₹)' column in one of the files.")
            continue

        # Common GSTINs
        common_gstin = df_company.index.intersection(df_payments.index)

        taxable_comparison = pd.DataFrame({
            'Company_Taxable_Value': df_company.loc[common_gstin, 'Taxable Value (₹)'],
            'Payment_Taxable_Value': df_payments.loc[common_gstin, 'Taxable Value (₹)']
        })

        taxable_comparison['Difference'] = taxable_comparison['Company_Taxable_Value'] - taxable_comparison['Payment_Taxable_Value']
        taxable_comparison['Claim Status'] = taxable_comparison['Difference'].apply(lambda x: 'Claim GST' if x == 0 else 'Claim Next Month')

        # Filter options
        filter_option = st.radio("📌 Show GSTINs with:", ["All", "Only Differences", "Only No Differences"], horizontal=True)

        if filter_option == "Only Differences":
            filtered_df = taxable_comparison[taxable_comparison['Difference'] != 0]
        elif filter_option == "Only No Differences":
            filtered_df = taxable_comparison[taxable_comparison['Difference'] == 0]
        else:
            filtered_df = taxable_comparison

        st.markdown("### 🔍 GSTIN Comparison Result")
        st.dataframe(filtered_df, use_container_width=True)

        # Records in Payments but not in Company
        unmatched_gstin = df_payments.index.difference(df_company.index)
        unmatched_records = df_payments.loc[unmatched_gstin]

        st.markdown("### ⚠️ Records Present in Payments but Missing in Company")
        st.dataframe(unmatched_records, use_container_width=True)

        # Email Section
        st.markdown("---")
        st.markdown("### 📧 Send Emails to Unmatched Suppliers")

        sender_email = st.text_input("Enter Sender Gmail Address")
        sender_password = st.text_input("Enter Sender Gmail App Password", type="password")

        if st.button("Send Emails"):
            if not sender_email or not sender_password:
                st.error("Sender email and password must be provided.")
            else:
                try:
                    yag = yagmail.SMTP(user=sender_email, password=sender_password)
                    success_count = 0

                    for gstin in unmatched_records.index:
                        row = unmatched_records.loc[gstin]
                        email = row.get('email')
                        name = row.get('Trade/Legal name', 'Supplier')

                        if pd.notnull(email) and isinstance(email, str):
                            subject = f"Missing Invoice Alert for GSTIN {gstin}"
                            body = f"""
Dear {name},

Our system shows that your invoice with GSTIN {gstin} is present in our payment records
but missing from our internal filing system.

Kindly share the corresponding invoice/documents at the earliest so we can reconcile records.

Regards,
Finance Team
"""
                            try:
                                yag.send(to=email, subject=subject, contents=body)
                                success_count += 1
                            except Exception as e:
                                st.warning(f"Failed to send email to {email}: {e}")

                    st.success(f"✅ Emails sent successfully to {success_count} unmatched suppliers.")

                except Exception as e:
                    st.error(f"❌ Failed to authenticate Gmail SMTP: {e}")
