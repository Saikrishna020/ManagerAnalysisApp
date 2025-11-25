import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Manager Case Analysis", layout="centered")

st.title("📊 Manager Case Analysis Tool")
st.write("Upload your Excel file and select the type of manager analysis you'd like to perform.")

# Upload the Excel file
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

# Let the user choose the type of analysis
analysis_type = st.selectbox("Select Analysis Type", [
    "Report Manager", "Assigning Manager", "Allotment Manager"
])

# Process once both file and selection are available
if uploaded_file and analysis_type:
    try:
        df = pd.read_excel(uploaded_file, skiprows=1)
        df.columns = df.columns.str.strip()  # Clean column names

        # Fill missing values
        if analysis_type == "Report Manager":
            df["Report Manager"].fillna(df["Manager"], inplace=True)
            report_col = "Report Manager"
            output_file = "manager_case_analysis_report.xlsx"
        elif analysis_type == "Assigning Manager":
            df["Assigning Manager"].fillna(df["Manager"], inplace=True)
            report_col = "Assigning Manager"
            output_file = "manager_case_analysis_assigning.xlsx"
        else:
            df["Allotment Manager"].fillna(df["Manager"], inplace=True)
            report_col = "Allotment Manager"
            output_file = "manager_case_analysis_allotment.xlsx"

        # Manager counts
        manager_cases = df["Manager"].value_counts().reset_index()
        manager_cases.columns = ["Manager", "Number of Cases"]

        # Report/Assigning/Allotment Manager counts
        other_cases = df[report_col].value_counts().reset_index()
        other_cases.columns = [report_col, "Number of Cases"]

        # Display results
        st.subheader("🔹 Manager Cases")
        st.dataframe(manager_cases)

        st.subheader(f"🔹 {report_col} Cases")
        st.dataframe(other_cases)

        # Export to Excel
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            manager_cases.to_excel(writer, sheet_name="Manager Cases", index=False)
            other_cases.to_excel(writer, sheet_name=f"{report_col} Cases", index=False)

        with open(output_file, "rb") as f:
            st.download_button(
                label="📥 Download Analysis Report",
                data=f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Something went wrong: {e}")
