import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.offline as pyo
import os
import glob
from datetime import datetime

st.set_page_config(page_title="Defect Dashboard", layout="wide")

def load_excel_data(file_path):
    try:
        try:
            df = pd.read_excel(file_path, sheet_name='Defects')
        except:
            df = pd.read_excel(file_path, sheet_name=0)
        
        df_mapped = pd.DataFrame({
            'dateAdded': df.iloc[:, 0] if len(df.columns) > 0 else None,
            'jiraId': df.iloc[:, 2] if len(df.columns) > 2 else None,
            'description': df.iloc[:, 3] if len(df.columns) > 3 else None,
            'priority': df.iloc[:, 4] if len(df.columns) > 4 else None,
            'dueDate': df.iloc[:, 5] if len(df.columns) > 5 else None,
            'status': df.iloc[:, 7] if len(df.columns) > 7 else None,
            'area': df.iloc[:, 9] if len(df.columns) > 9 else None,
            'category': df.iloc[:, 12] if len(df.columns) > 12 else None,
            'releaseDate': df.iloc[:, 15] if len(df.columns) > 15 else None,
            'passOrFail': df.iloc[:, 19] if len(df.columns) > 19 else None,
            'partyAssigned': df.iloc[:, 23] if len(df.columns) > 23 else None
        })
        
        df_mapped = df_mapped.dropna(subset=['jiraId', 'description'], how='all')
        return df_mapped
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return None

st.title("Defects Status")

# File uploader
uploaded_file = st.file_uploader("Upload Excel file", type=['xlsx', 'xls'])

if uploaded_file is not None:
    st.info(f"Loading: {uploaded_file.name}")
    data = load_excel_data(uploaded_file)
else:
    st.warning("Please upload an Excel file to continue")
    st.stop()

if data is None:
    st.stop()

st.success(f"Loaded {len(data)} defects from {os.path.basename(file_path)}")

# Metrics
wip = len(data[data['status'] == 'WIP'])
p1p2_open = len(data[(data['priority'].isin([1, 2])) & (data['status'] == 'WIP')])
wip_no_release = len(data[(data['status'] == 'WIP') & (data['releaseDate'].isna())])
product_limitation = len(data[data['status'] == 'Product Limitation'])

col1, col2, col3, col4 = st.columns(4)
col1.metric("Open", wip)
col2.metric("P1/P2 Open", p1p2_open)
col3.metric("No Release Date", wip_no_release)
col4.metric("Product Limitations", product_limitation)

# Charts
col1, col2 = st.columns(2)

# Store chart data for export
wip_data = data[data['status'] == 'WIP']
priority_counts = wip_data['priority'].value_counts().sort_index()
priority_counts.index = [f"P{int(p)}" for p in priority_counts.index]

status_counts = data['status'].value_counts()
status_order = ['WIP', 'CLOSED', 'Product Limitation']
status_labels = {'WIP': 'Open', 'CLOSED': 'Closed', 'Product Limitation': 'Product Limitation'}
ordered_status = []
ordered_values = []

for status in status_order:
    if status in status_counts.index:
        ordered_status.append(status_labels.get(status, status))
        ordered_values.append(status_counts[status])

for status in status_counts.index:
    if status not in status_order:
        ordered_status.append(status)
        ordered_values.append(status_counts[status])

with col1:
    st.subheader("Priority Distribution")
    fig = px.bar(x=priority_counts.index, y=priority_counts.values, 
                 labels={'x': 'Priority', 'y': 'Count'})
    st.plotly_chart(fig, use_container_width=True)

with col2:
    st.subheader("Status Distribution")
    fig = go.Figure(data=[go.Pie(
        labels=ordered_status, 
        values=ordered_values,
        textinfo='label',
        hovertemplate='<b>%{label}</b><br>Count: %{value}<extra></extra>'
    )])
    fig.update_layout(showlegend=True, legend=dict(traceorder='normal'))
    fig.update_traces(sort=False)
    st.plotly_chart(fig, use_container_width=True)

# Table
wip_no_release_data = data[(data['status'] == 'WIP') & (data['releaseDate'].isna())]
if not wip_no_release_data.empty:
    st.subheader("Open Defects with No Release Date")
    display_cols = ['jiraId', 'description', 'priority', 'area', 'partyAssigned']
    table_data = wip_no_release_data[display_cols].copy()
    table_data['priority'] = table_data['priority'].astype(int)
    table_data.columns = ['JIRA ID', 'Description', 'Priority', 'Area', 'Assigned to']
    st.dataframe(table_data, use_container_width=True, hide_index=True)

# Auto-generate HTML export
current_datetime = datetime.now().strftime("%d/%m/%Y at %I:%M %p")
table_html = ""
if not wip_no_release_data.empty:
    table_html = f"""
    <h3>Open Defects with No Release Date</h3>
    {table_data.to_html(index=False, escape=False, table_id='defects-table')}
    """

html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <title>Defects Dashboard</title>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
        * {{ box-sizing: border-box; }}
        body {{ 
            font-family: 'Source Sans Pro', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            margin: 0;
            padding: 2rem;
            background-color: #ffffff;
            color: #262730;
            line-height: 1.6;
        }}
        .container {{ max-width: 1200px; margin: 0 auto; }}
        h1 {{ 
            color: #262730;
            font-size: 2.5rem;
            font-weight: 600;
            margin-bottom: 1rem;
            border-bottom: 2px solid #f0f2f6;
            padding-bottom: 1rem;
        }}
        h3 {{ 
            color: #262730;
            font-size: 1.25rem;
            font-weight: 600;
            margin: 1.5rem 0 1rem 0;
        }}
        .metrics {{ 
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 1.5rem;
            margin: 2rem 0;
        }}
        .metric {{ 
            background: #ffffff;
            border: 1px solid #e6eaf1;
            border-radius: 0.5rem;
            padding: 1.5rem;
            text-align: center;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }}
        .metric-value {{ 
            font-size: 2.5rem;
            font-weight: 700;
            color: #ff4b4b;
            margin-bottom: 0.5rem;
            line-height: 1;
        }}
        .metric-label {{ 
            font-size: 0.875rem;
            color: #8e9297;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }}
        table {{ 
            width: 100%;
            border-collapse: collapse;
            margin: 2rem 0;
            background: #ffffff;
            border-radius: 0.5rem;
            overflow: hidden;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }}
        th {{ 
            background: #f8f9fa;
            color: #262730;
            font-weight: 600;
            padding: 1rem;
            text-align: left;
            border-bottom: 2px solid #e6eaf1;
            font-size: 0.875rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }}
        td {{ 
            padding: 0.75rem 1rem;
            border-bottom: 1px solid #f0f2f6;
            color: #262730;
        }}
        tr:hover {{ background-color: #f8f9fa; }}
        .table-container {{ 
            background: #ffffff;
            border: 1px solid #e6eaf1;
            border-radius: 0.5rem;
            overflow: hidden;
            margin: 2rem 0;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>Defects Status</h1>
        <p style="color: #8e9297; font-size: 0.875rem; margin-bottom: 2rem;">Generated on {current_datetime}</p>
        
        <div class="metrics">
            <div class="metric">
                <div class="metric-value">{wip}</div>
                <div class="metric-label">Open</div>
            </div>
            <div class="metric">
                <div class="metric-value">{p1p2_open}</div>
                <div class="metric-label">P1/P2 Open</div>
            </div>
            <div class="metric">
                <div class="metric-value">{wip_no_release}</div>
                <div class="metric-label">No Release Date</div>
            </div>
            <div class="metric">
                <div class="metric-value">{product_limitation}</div>
                <div class="metric-label">Product Limitations</div>
            </div>
        </div>
        
        <div class="table-container">
            {table_html}
        </div>
    </div>
</body>
</html>
"""

# Try to save to original folder, fallback to current directory
export_folder = r"C:\Users\danjw\Greater Wellington Regional Council\[Ext] Metlink RTI 2.0 - User Acceptance Testing"
if not os.path.exists(export_folder):
    export_folder = os.getcwd()
    
export_file_path = os.path.join(export_folder, "Defects Dashboard.html")

try:
    with open(export_file_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    st.success(f"HTML dashboard automatically saved to: {export_file_path}")
except Exception as e:
    st.error(f"Failed to save file: {e}")

st.download_button(
    label="Download Defects Dashboard",
    data=html_content,
    file_name="Defects Dashboard.html",
    mime="text/html"
)