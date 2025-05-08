import streamlit as st
import pandas as pd
import yagmail
from io import BytesIO

st.set_page_config(page_title="GST Comparison App", layout="wide")

st.title("📑 GSTIN-Based Excel Comparison & Email Automation")

# Upload Excel files
company_file = st.file_uploader("Upload Company Excel File", type=["xlsx"], key="company")
payments_file = st.file_uploader("Upload Payments Excel File", type=["xlsx"], key="payments")

# Email credentials
sender_email = st.text_input("Enter Sender Gmail ID")
sender_password = st.text_input("Enter Sender Gmail Password", type="password")

if company_file and payments_file:
    # Load Excel files with all sheets
    company_df_all = pd.read_excel(company_file, sheet_name=None, index_col='GSTIN of supplier')
    payments_df_all = pd.read_excel(payments_file, sheet_name=None, index_col='GSTIN of supplier')

    common_sheets = set(company_df_all.keys()).intersection(set(payments_df_all.keys()))

    for sheet in common_sheets:
        st.markdown(f"## 📄 Sheet: {sheet}")

        df_company = company_df_all[sheet]
        df_payments = payments_df_all[sheet]

        if 'Taxable Value (₹)' not in df_company.columns or 'Taxable Value (₹)' not in df_payments.columns:
            st.warning(f"Missing 'Taxable Value (₹)' in sheet {sheet}")
            continue

        # Align by GSTIN
        common_gstin = df_company.index.intersection(df_payments.index)

        taxable_comparison = pd.DataFrame({
            'Company_Taxable_Value': df_company.loc[common_gstin, 'Taxable Value (₹)'],
            'Payment_Taxable_Value': df_payments.loc[common_gstin, 'Taxable Value (₹)']
        })

        # Compute difference
        taxable_comparison['Difference'] = (
            taxable_comparison['Company_Taxable_Value'] - taxable_comparison['Payment_Taxable_Value']
        )

        # Add Claim Status
        taxable_comparison['Claim Status'] = taxable_comparison['Difference'].apply(
            lambda x: 'Claim GST' if x == 0 else 'Claim Next Month'
        )

        # Filter options
        filter_option = st.radio(
            f"🎯 Filter GSTINs in {sheet}:",
            ('Show All', 'Only Claim GST', 'Only Claim Next Month'),
            key=f'filter_{sheet}'
        )

        if filter_option == 'Only Claim GST':
            filtered_df = taxable_comparison[taxable_comparison['Claim Status'] == 'Claim GST']
        elif filter_option == 'Only Claim Next Month':
            filtered_df = taxable_comparison[taxable_comparison['Claim Status'] == 'Claim Next Month']
        else:
            filtered_df = taxable_comparison

        # Conditional styling
        def highlight_status(row):
            color = 'lightgreen' if row['Claim Status'] == 'Claim GST' else 'lightyellow'
            return ['background-color: {}'.format(color) if col == 'Claim Status' else '' for col in row.index]

        st.markdown("### 🔍 GSTIN Comparison Result")
        st.dataframe(filtered_df.style.apply(highlight_status, axis=1), use_container_width=True)

        # Records in payments but not in company
        unmatched_gstin = df_payments.index.difference(df_company.index)
        unmatched_records = df_payments.loc[unmatched_gstin]

        st.markdown("### 🚨 Invoices in Payments Sheet but Missing in Company Sheet")
        st.dataframe(unmatched_records)

        # Button to trigger email sending
        if sender_email and sender_password:
            if st.button(f"📧 Send Emails for Missing Records in {sheet}", key=f'email_{sheet}'):
                try:
                    yag = yagmail.SMTP(user=sender_email, password=sender_password)
                    sent = 0
                    for gstin in unmatched_records.index:
                        row = unmatched_records.loc[gstin]
                        email = row.get('email')
                        name = row.get('Trade/Legal name', 'Supplier')

                        if pd.notnull(email) and isinstance(email, str):
                            subject = f"Missing Invoice Alert for GSTIN {gstin}"
                            body = f"""
Dear {name},

Our system shows that your invoice with GSTIN {gstin} is present in our payment records but missing from our internal records.

Kindly share the corresponding invoice/documents at the earliest so we can reconcile.

Regards,  
Finance Team
"""
                            yag.send(to=email, subject=subject, contents=body)
                            sent += 1
                    st.success(f"✅ {sent} email(s) sent successfully.")
                except Exception as e:
                    st.error(f"❌ Email sending failed: {e}")
        else:
            st.warning("Please provide Gmail ID and Password to send emails.")

