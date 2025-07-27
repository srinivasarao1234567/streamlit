import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Internship Dashboard", layout="wide")

# Load Excel
xls = pd.ExcelFile("marksexcel.xlsx")        
df1 = xls.parse(xls.sheet_names[0], skiprows=3)
df2 = xls.parse(xls.sheet_names[1], skiprows=3)

# Rename and clean columns
df1.columns = ["RegNo", "StudentName", "Organization", "Course", "Unused1", "Marks"]
df2.columns = ["RegNo", "StudentName", "Organization", "Course", "Unused1", "Marks"]
df1["Batch"] = "Y21"
df2["Batch"] = "Y22"

df = pd.concat([df1, df2], ignore_index=True)
df = df[["RegNo", "StudentName", "Organization", "Course", "Marks", "Batch"]]
df["Marks"] = pd.to_numeric(df["Marks"], errors="coerce")
df.dropna(subset=["Organization", "Course", "Marks"], inplace=True)

# Sidebar
st.sidebar.title("🔧 Filter Options")

# Batch selector
selected_batch = st.sidebar.selectbox("Select Batch", ["All", "Y21", "Y22"])

# Dynamic slider limits
unique_org_count = df["Organization"].nunique()
unique_course_count = df["Course"].nunique()

top_n_org = st.sidebar.slider("Top Organizations", min_value=3, max_value=unique_org_count, value=5)
top_n_course = st.sidebar.slider("Top Courses", min_value=3, max_value=unique_course_count, value=5)

# Histogram settings
hist_bins = st.sidebar.slider("Histogram Bins", min_value=5, max_value=50, value=7)
bin_color = st.sidebar.color_picker("Pick Histogram Bin Color", value="#1f77b4")

# Sidebar - Project members
with st.sidebar.expander("👨‍💻 Project Members List", expanded=True):
    st.write(""" 
- R. Moditha - L23AIT542  
- M. Sri Vani Prasanna - Y22AIT470  
- M. Aswitha - Y22AIT477 
- N. Komali Kiran - Y22AIT483  
- Sk. Vaheeda - Y22AIT510
    """)

# Filter batch
if selected_batch != "All":
    df = df[df["Batch"] == selected_batch]

# Tabs for charts
tab1, tab2, tab3 = st.tabs(["📊 Top Organizations", "📈 Marks Histogram", "📘 Top Courses"])

# PIE CHART (Plotly): Top Organizations
with tab1:
    org_counts = df["Organization"].value_counts().nlargest(top_n_org).reset_index()
    org_counts.columns = ['Organization', 'Count']
    fig1 = px.pie(
        org_counts,
        names='Organization',
        values='Count'
    )
    fig1.update_traces(
        textinfo='none',
        hovertemplate='%{label}: %{percent}<extra></extra>'
    )
    fig1.update_layout(title='Top Organizations by Count')
    st.plotly_chart(fig1, use_container_width=True)

# HISTOGRAM (Plotly): Marks Distribution
with tab2:
    fig_hist = go.Figure()

    counts, bins = pd.cut(df["Marks"], bins=hist_bins, retbins=True, right=False)
    bin_df = df.groupby(counts)["Marks"].count().reset_index(name="Count")
    total = bin_df["Count"].sum()
    bin_labels = [f"{interval.left:.0f}-{interval.right:.0f}" for interval in bin_df["Marks"]]
    percentages = (bin_df["Count"] / total * 100).round(1)

    fig_hist.add_trace(go.Bar(
        x=bin_labels,
        y=bin_df["Count"],
        marker_color=bin_color,
        hovertemplate='%{x} Marks: %{y} students<br>Percentage: %{customdata:.1f}%<extra></extra>',
        customdata=percentages
    ))

    fig_hist.update_layout(
        title="Marks Distribution Histogram",
        xaxis_title="Marks Range",
        yaxis_title="Number of Students",
        bargap=0.1
    )
    st.plotly_chart(fig_hist, use_container_width=True)

# PIE CHART (Plotly): Top Courses
with tab3:
    course_counts = df["Course"].value_counts().nlargest(top_n_course).reset_index()
    course_counts.columns = ['Course', 'Count']
    fig2 = px.pie(
        course_counts,
        names='Course',
        values='Count'
    )
    fig2.update_traces(
        textinfo='none',
        hovertemplate='%{label}: %{percent}<extra></extra>'
    )
    fig2.update_layout(title='Top Courses by Count')
    st.plotly_chart(fig2, use_container_width=True)
