import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image
from datetime import datetime
import os
import warnings
import time

warnings.filterwarnings("ignore")

# ---------------- CONFIG ----------------
st.set_page_config(page_title="Secure Excel Dashboard", layout="wide")

# ---------------- HEADER ----------------
logo = Image.open("static/oma-logo.jpg")
col1, col2 = st.columns([0.2, 0.8])
with col1:
    st.image(logo, width=140)
with col2:
    st.markdown(
        """
        <h2 style="
            font-weight:bold;
            padding:6px;
            border-radius:6px;
            background-color:orange;
            display:inline-block;
            text-align:center;
        ">
            OMA-HO Sheet - UAE Region 2025
        </h2>
        """,
        unsafe_allow_html=True
    )
st.title("🏪 WELCOME TO OMA HARDWARE TEAM - HO (UAE)")

# ---------------- FOLDERS ----------------
EXCEL_FOLDER = "excel_files"
LOG_FOLDER = "logs"
LOG_FILE = os.path.join(LOG_FOLDER, "login_activity.csv")

# ---------------- USERS ----------------
USERS = {
    "admin": {"password": "admin", "role": "Admin"},
    "manager": {"password": "manager123", "role": "Manager"}
}

# ---------------- LOG FILE INIT ----------------
if not os.path.exists(LOG_FOLDER):
    os.makedirs(LOG_FOLDER)
if not os.path.exists(LOG_FILE):
    pd.DataFrame(columns=["Username", "Role", "Action", "Timestamp"]).to_csv(LOG_FILE, index=False)

# ---------------- SESSION STATE ----------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user = None
    st.session_state.role = None

# ---------------- LOG FUNCTION ----------------
def log_activity(username, role, action):
    log_df = pd.read_csv(LOG_FILE)
    new_entry = {
        "Username": username,
        "Role": role,
        "Action": action,
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    log_df = pd.concat([log_df, pd.DataFrame([new_entry])], ignore_index=True)
    log_df.to_csv(LOG_FILE, index=False)

# ---------------- LOGIN CALLBACK ----------------
def login_callback(username, password):
    if username in USERS and USERS[username]["password"] == password:
        st.session_state.logged_in = True
        st.session_state.user = username
        st.session_state.role = USERS[username]["role"]
        log_activity(username, USERS[username]["role"], "LOGIN")
    else:
        st.error("❌ Invalid username or password")

# ---------------- LOGOUT CALLBACK ----------------
def logout_callback():
    log_activity(st.session_state.user, st.session_state.role, "LOGOUT")
    st.session_state.logged_in = False
    st.session_state.user = None
    st.session_state.role = None

# ---------------- LOGIN PAGE ----------------
if not st.session_state.logged_in:
    st.title("🔐 Secure Login")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        login_callback(username, password)
    st.stop()

# ---------------- SIDEBAR ----------------
st.sidebar.success(f"👤 {st.session_state.user} ({st.session_state.role})")
st.sidebar.button("🚪 Logout", on_click=logout_callback)
if st.session_state.role == "Admin":
    with st.sidebar.expander("📜 Login Activity Log"):
        st.dataframe(pd.read_csv(LOG_FILE), use_container_width=True)

# ---------------- LOAD EXCEL ----------------
if not os.path.exists(EXCEL_FOLDER):
    st.error("❌ Excel folder not found")
    st.stop()

excel_files = [f for f in os.listdir(EXCEL_FOLDER) if f.endswith((".xlsx", ".xls"))]
if not excel_files:
    st.warning("⚠️ No Excel files found")
    st.stop()

selected_file = st.sidebar.radio("📂 Select Excel File", excel_files)
file_path = os.path.join(EXCEL_FOLDER, selected_file)
xls = pd.ExcelFile(file_path)
sheet = st.sidebar.selectbox("📑 Select Sheet", xls.sheet_names)
df_original = xls.parse(sheet)

# ---------------- HANDLE DUPLICATE COLUMNS ----------------
cols = pd.Series(df_original.columns)
for dup in cols[cols.duplicated()].unique():
    cols[cols[cols == dup].index[1:]] = [f"{dup}_{i}" for i in range(1, sum(cols == dup))]
df_original.columns = cols

st.subheader(f"📄 Data Preview — {selected_file} / {sheet}")
st.dataframe(df_original, use_container_width=True)

# ---------------- DEPARTMENT & NUMERIC ----------------
dept_keywords = ["department", "dept", "division", "section", "team", "unit"]
dept_cols = [c for c in df_original.columns if any(k in c.lower() for k in dept_keywords)]
dept_col = dept_cols[0] if dept_cols else st.selectbox("Select Department Column", df_original.columns)
numeric_cols = df_original.select_dtypes(include="number").columns.tolist()
if not numeric_cols:
    st.warning("⚠️ No numeric columns available")
    st.stop()

# ---------------- DEPARTMENT FILTER ----------------
departments = df_original[dept_col].dropna().unique().tolist()
selected_dept = st.sidebar.selectbox("🏢 Select Department", ["All"] + departments)
df = df_original.copy()
if selected_dept != "All":
    df = df[df[dept_col] == selected_dept]

# ---------------- KPIs ----------------
st.markdown("## 📊 KPIs")
kpi_col = st.selectbox("Select KPI Metric", numeric_cols)
c1, c2, c3, c4 = st.columns(4)
c1.metric("Total", f"{df[kpi_col].sum():,.2f}")
c2.metric("Average", f"{df[kpi_col].mean():,.2f}")
c3.metric("Max", f"{df[kpi_col].max():,.2f}")
c4.metric("Min", f"{df[kpi_col].min():,.2f}")

# ---------------- GROWTH ----------------
st.markdown("## 📈 Growth")
if len(df) > 1:
    first_val = df[kpi_col].iloc[0]
    last_val = df[kpi_col].iloc[-1]
    if first_val != 0:
        growth = ((last_val - first_val)/first_val)*100
        st.metric("Growth %", f"{growth:.2f}%")
    else:
        st.info("Growth % not calculable (first value = 0)")
else:
    st.info("Not enough data for growth calculation")

# ---------------- CHARTS ----------------
st.markdown("## 🧠 Department Charts")
x_col = st.selectbox("X-Axis", df.columns)
y_col = st.selectbox("Y-Axis (Numeric)", numeric_cols)
auto_mode = st.checkbox("Enable Smart Chart Mode", True)

def is_date(series):
    try:
        pd.to_datetime(series)
        return True
    except:
        return False

if auto_mode:
    chart_type = "Line" if is_date(df[x_col]) else "Pie" if df[x_col].nunique() <= 8 else "Bar"
else:
    chart_type = st.selectbox("Chart Type", ["Bar", "Line", "Pie"])

cols = st.columns(3)
with cols[0]:
    fig_bar = px.bar(df, x=x_col, y=y_col, template="plotly_dark", height=500)
    st.plotly_chart(fig_bar, use_container_width=True)
    st.caption("📊 Bar Chart")

with cols[1]:
    fig_line = px.line(df, x=x_col, y=y_col, markers=True, template="plotly_dark", height=500)
    st.plotly_chart(fig_line, use_container_width=True)
    st.caption("📈 Line Chart")

with cols[2]:
    fig_pie = px.pie(df, names=x_col, values=y_col, template="plotly_dark", height=500)
    st.plotly_chart(fig_pie, use_container_width=True)
    st.caption("🥧 Pie Chart")

# ---------------- TREND GRAPH ----------------
st.markdown("## 📊 Trend Graph")
trend_x = st.selectbox("Select Trend X-Axis", df.columns, key="trend_x")
trend_y = st.selectbox("Select Trend Y-Axis", numeric_cols, key="trend_y")

fig_trend = go.Figure()
fig_trend.add_trace(go.Scatter(
    x=df[trend_x],
    y=df[trend_y],
    mode='lines+markers',
    marker=dict(color='orange', size=8),
    line=dict(width=3)
))
fig_trend.update_layout(
    title=f"Trend of {trend_y} over {trend_x}",
    template="plotly_dark",
    height=500
)
st.plotly_chart(fig_trend, use_container_width=True)

# ---------------- DEPARTMENT SUMMARY ----------------
dept_summary = df_original.groupby(dept_col, as_index=False)[kpi_col].sum()
# Ensure unique column names
cols_summary = pd.Series(dept_summary.columns)
for dup in cols_summary[cols_summary.duplicated()].unique(): 
    cols_summary[cols_summary[cols_summary == dup].index[1:]] = [f"{dup}_{i}" for i in range(1, sum(cols_summary == dup))]
dept_summary.columns = cols_summary
dept_summary = dept_summary.sort_values(by=kpi_col, ascending=False)

# ---------------- TOP & BOTTOM ----------------
st.markdown("## 🏆 Top & Bottom Departments")
if not dept_summary.empty:
    top_dept = dept_summary.iloc[0]
    bottom_dept = dept_summary.iloc[-1]
    t1, t2 = st.columns(2)
    t1.metric("Top Department", top_dept[dept_col], f"{top_dept[kpi_col]:,.2f}")
    t2.metric("Bottom Department", bottom_dept[dept_col], f"{bottom_dept[kpi_col]:,.2f}")
else:
    st.info("⚠️ No department data for Top/Bottom metrics.")

# ---------------- DEPARTMENT SUMMARY TABLE ----------------
st.markdown("## 🧾 Department Summary Table")
if not dept_summary.empty:
    st.dataframe(dept_summary, use_container_width=True)
else:
    st.info("No data available")

# ---------------- DOWNLOAD ----------------
st.download_button(
    "📥 Download Chart (HTML)",
    fig_bar.to_html(),
    file_name=f"{selected_file}_{selected_dept}_charts.html",
    mime="text/html"
)

# ---------------- ANIMATED DONUT ----------------
st.markdown("## 🍩 Animated Rotating Donut Chart")
donut_name_col = st.selectbox("Donut: Name Column", df.columns, key="donut_name")
donut_value_col = st.selectbox("Donut: Value Column", numeric_cols, key="donut_value")

# Aggregate small slices safely
df_donut = df.groupby(donut_name_col, as_index=False)[donut_value_col].sum()
# Combine small values into 'Other'
threshold = df_donut[donut_value_col].sum() * 0.05
df_donut["Category"] = df_donut.apply(
    lambda row: row[donut_name_col] if row[donut_value_col] >= threshold else "Other", axis=1
)
df_donut = df_donut.groupby("Category", as_index=False)[donut_value_col].sum()

# Plot donut
fig_donut = go.Figure(
    go.Pie(
        labels=df_donut["Category"],
        values=df_donut[donut_value_col],
        hole=0.5,
        pull=[0.1 if i == df_donut[donut_value_col].idxmax() else 0 for i in range(len(df_donut))],
        textinfo="label+percent",
        sort=False
    )
)

# ---------------- DISPLAY DONUT ----------------
st.plotly_chart(fig_donut, use_container_width=True)

st.markdown(" 🔧🔧🔧🔧For support contact = ankush@omaemirates.com ")
time.sleep(0.25)