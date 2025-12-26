import streamlit as st
import pandas as pd
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(page_title="Agent Performance Processor", layout="wide")

st.title("📊 Agent Performance Data Processor")

# ------------------ FUNCTIONS (UNCHANGED LOGIC) ------------------

def load_and_clean_data(uploaded_file):
    lines = uploaded_file.getvalue().decode("utf-8").splitlines(True)

    header_row = 0
    for i, line in enumerate(lines):
        if 'USER NAME' in line.upper():
            header_row = i
            break

    metadata_rows = lines[:header_row]

    df = pd.read_csv(
        uploaded_file,
        skiprows=header_row,
        on_bad_lines='skip',
        engine='python'
    )

    columns_to_delete = [
        'CURRENT USER GROUP','MOST RECENT USER GROUP','PAUSAVG','WAITAVG',
        'TALKAVG','DISPAVG','DEADAVG','CUSTAVG','ANS','SSMS','REDIAL',
        'test','testne','TestIT','TESTNC','TESTCB','Test22','DUPLICATE CALLS'
    ]

    df = df.drop(columns=[c for c in columns_to_delete if c in df.columns])

    if len(df) > 0:
        df = df.iloc[:-1]

    return df, metadata_rows


def process_time_columns(df):
    df['TOTAL PAUSE'] = (
        pd.to_timedelta(df['PAUSE'], errors='coerce') +
        pd.to_timedelta(df['DEAD'], errors='coerce') +
        pd.to_timedelta(df['DISPO'], errors='coerce')
    )

    df['TOTAL PAUSE'] = df['TOTAL PAUSE'].apply(
        lambda x: f"{int(x.total_seconds()//3600):02d}:"
                  f"{int((x.total_seconds()%3600)//60):02d}:"
                  f"{int(x.total_seconds()%60):02d}"
        if pd.notna(x) else "00:00:00"
    )
    return df


def reorder_and_sort(df):
    df['ID'] = pd.to_numeric(df['ID'], errors='coerce').fillna(0).astype(int)
    df = df.sort_values(by='TOTAL INBOUND CALLS', ascending=False)
    df.reset_index(drop=True, inplace=True)
    df.index += 1

    desired_columns = [
        'USER NAME','ID','CALLS','TIME','PAUSE','WAIT','TALK',
        'DISPO','DEAD','TOTAL PAUSE','CUSTOMER',
        'TOTAL INBOUND CALLS','TOTAL OUTBOUND CALLS'
    ]

    df = df[[c for c in desired_columns if c in df.columns]]
    df['REMARKS'] = ''
    return df


def save_to_excel(df, metadata_rows, file_path):
    from openpyxl import load_workbook
    df.to_excel(file_path, index=False)

# ------------------ STREAMLIT UI ------------------

uploaded_file = st.file_uploader("📂 Upload Agent Performance CSV", type=["csv"])

if uploaded_file:
    st.success("File uploaded successfully!")

    df, metadata = load_and_clean_data(uploaded_file)
    df = process_time_columns(df)
    df = reorder_and_sort(df)

    st.subheader("🔍 Preview Processed Data")
    st.dataframe(df.head(10), width=1200)

    excel_file = "styled_agent_performance.xlsx"
    save_to_excel(df, metadata, excel_file)

    with open(excel_file, "rb") as f:
        st.download_button(
            "⬇ Download Excel Report",
            f,
            file_name=excel_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.subheader("📈 Summary")
    st.metric("Total Agents", len(df))
    st.metric("Total Inbound Calls", int(df['TOTAL INBOUND CALLS'].sum()))
    st.metric("Average Inbound Calls", round(df['TOTAL INBOUND CALLS'].mean(), 2))