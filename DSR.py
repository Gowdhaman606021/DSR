import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from scipy.interpolate import make_interp_spline
import os
from datetime import datetime
import plotly.express as px
import streamlit.components.v1 as components

# Get current date
today = datetime.today().strftime('%B %d, %Y')  # e.g., September 14, 2025



# Set page config
st.set_page_config(layout="wide")

# Display header with date
st.header(f"Daily Status Report ({today})")

st.markdown("---")

df1 = pd.read_excel("DSR_Sample.xlsx", sheet_name=0, engine='openpyxl')

if df1.shape[1] >= 3:
    color_map = {
        1: "#BDE3C3",  # Green
        2: "#FFE797",  # Amber
        3: "#FFC7A7"   # Red
    }

    cols = st.columns(len(df1))
    for idx, (index, row) in enumerate(df1.iterrows()):
        label = str(row.iloc[0])
        status = int(row.iloc[2])
        color = color_map.get(status, "#e2e3e5")  # Default gray

        with cols[idx % len(cols)]:
            st.markdown(f"""
                <div style='background-color:{color}; padding:10px; border-radius:8px; text-align:center; font-weight:bold;'>
                    {label}
                </div>
            """, unsafe_allow_html=True)
else:
    st.warning("Sheet1 must contain at least three columns.")

st.markdown("---")

# Load Excel file
file_path = "DSR_Sample.xlsx"
if os.path.exists(file_path):
    # Load all sheets
    df1 = pd.read_excel(file_path, sheet_name=0, engine='openpyxl')
    df2 = pd.read_excel(file_path, sheet_name=1, engine='openpyxl')
    df3 = pd.read_excel(file_path, sheet_name=2, engine='openpyxl')
    df5 = pd.read_excel(file_path, sheet_name=4, engine='openpyxl')
    df3.columns = [col.strip() for col in df3.columns]

    # --- Tile View with Donut Charts ---
    if df1.shape[1] >= 4:
        if "selected_tile" not in st.session_state:
            st.session_state.selected_tile = None

        num_columns = 4
        rows = [df1.iloc[i:i+num_columns] for i in range(0, len(df1), num_columns)]

        for row in rows:
            cols = st.columns(num_columns)
            for idx, (index, data) in enumerate(row.iterrows()):
                with cols[idx]:
                    tile_label = str(data[0])
                    status_value = int(data[2])
                    completion = float(data[3])

                    color = {
                        1: "#f8d7da",
                        2: "#fff3cd",
                        3: "#d4edda"
                    }.get(status_value, "#e2e3e5")

                    st.markdown(f"""
                        <style>
                        div[data-testid="tile_{index}"] > button {{
                            background-color: {color};
                            color: black;
                            height: 100px;
                            width: 200px;
                            border-radius: 10px;
                            font-weight: bold;
                            margin: 5px;
                        }}
                        </style>
                    """, unsafe_allow_html=True)

                    if st.button(tile_label, key=f"tile_{index}"):
                        st.session_state.selected_tile = index

                    # Plotly donut chart
                    fig = go.Figure(data=[go.Pie(
                        values=[completion, 100 - completion],
                        hole=0.6,
                        marker=dict(colors=["#4CAF50", "#e0e0e0"]),
                        textinfo='none',
                        hoverinfo='label+percent'
                    )])
                    fig.update_layout(
                        showlegend=False,
                        margin=dict(t=0, b=0, l=0, r=0),
                        height=150,
                        width=150,
                        annotations=[dict(text=f"{completion:.0f}%", x=0.5, y=0.5, font_size=14, showarrow=False)],
                        template="plotly_dark"
                    )
                    st.plotly_chart(fig)

        if st.session_state.selected_tile is not None:
            selected_data = df1.iloc[st.session_state.selected_tile]
            st.markdown("---")
            st.subheader(f"Details for: {selected_data[0]}")
            st.write(selected_data[1])
    else:
        st.warning("Sheet1 must contain at least four columns.")

    # --- Sheet2 Table ---
    st.markdown("---")
    st.subheader("📋 Progress Table")
    st.dataframe(df2)


    # --- Sheet4 Table with Summary and Filter ---
    st.markdown("---")
    st.subheader("📄 Sign off Status")

    file_path = "DSR_Sample.xlsx"
    if os.path.exists(file_path):
        try:
            df4 = pd.read_excel(file_path, sheet_name=3, header=0, engine='openpyxl')
            df4 = df4.replace([pd.NA, None], "").dropna(how='all')

            if not df4.empty:
                status_col = df4.columns[2]
                status_counts = df4[status_col].value_counts()

                            # Filter
                status_options = ['All'] + sorted(df4[status_col].dropna().unique())
                selected_status = st.selectbox("Select status to filter:", status_options)

                filtered_df = df4 if selected_status == 'All' else df4[df4[status_col] == selected_status]


                # Highlighting
                def highlight_status(val):
                    val_str = str(val).strip().lower()
                    if val_str == 'signed off':
                        return 'background-color: #BDE3C3'
                    elif val_str == 'in progress':
                        return 'background-color: #FFE797'
                    elif val_str == 'blocked':
                        return 'background-color: #FFC7A7'
                    else:
                        return ''


                styled_df = filtered_df.style.applymap(highlight_status, subset=[status_col])
                st.dataframe(styled_df)
            else:
                st.info("Sheet4 is empty or contains no valid data.")
        except Exception as e:
            st.error(f"Error reading Sheet4: {e}")
    else:
        st.error("Excel file not found.")

    # --- Sheet2 Table ---
    st.markdown("---")
    st.subheader("📋 Open - Defect Table")
    st.dataframe(df5)

    # --- Sheet6 Donut Charts with Individual Totals ---
    st.markdown("---")

    file_path = "DSR_Sample.xlsx"
    if os.path.exists(file_path):
        try:
            # Read Sheet6 (index 5)
            df6 = pd.read_excel(file_path, sheet_name=5, header=0, engine='openpyxl')
            df6 = df6.replace([pd.NA, None], "").dropna(how='all')

            if not df6.empty:
                # Chart 1: Donut chart using Column 1 and Column 2
                col1, col2 = df6.columns[0], df6.columns[1]
                chart1_data = df6.groupby(col1)[col2].sum().reset_index()
                total1 = chart1_data[col2].sum()
                fig1 = px.pie(chart1_data, names=col1, values=col2,
                              title=f"Defect Distribution - Open Defects (Total: {total1})", hole=0.4)
                fig1.update_traces(textinfo='value')

                # Chart 2: Donut chart using Column 4 and Column 5
                col4, col5 = df6.columns[3], df6.columns[4]
                chart2_data = df6.groupby(col4)[col5].sum().reset_index()
                total2 = chart2_data[col5].sum()
                fig2 = px.pie(chart2_data, names=col4, values=col5,
                              title=f"Defect Distribution - Total Defects Raised (Total: {total2})", hole=0.4)
                fig2.update_traces(textinfo='value')

                # Display charts side by side
                col_left, col_right = st.columns(2)
                with col_left:
                    st.plotly_chart(fig1, use_container_width=True)
                with col_right:
                    st.plotly_chart(fig2, use_container_width=True)

            else:
                st.info("Sheet6 is empty or contains no valid data.")
        except Exception as e:
            st.error(f"Error reading Sheet6: {e}")
    else:
        st.error("Excel file not found.")

    # --- Combined Charts from Sheet3 ---
    st.markdown("---")
    st.subheader("📈 Progress Charts")

    # Chart 1: Burndown Chart (A to C)
    chart1 = None
    if df3.shape[1] >= 3:
        df_burn = df3.iloc[:, 0:3].copy()
        df_burn.columns = ['Date', 'Planned', 'Actual']
        df_burn['Date'] = pd.to_datetime(df_burn['Date'], errors='coerce')
        df_burn['Planned'] = pd.to_numeric(df_burn['Planned'], errors='coerce')
        df_burn['Actual'] = pd.to_numeric(df_burn['Actual'], errors='coerce')
        df_burn = df_burn.replace([np.inf, -np.inf], np.nan).dropna()

        if len(df_burn) >= 4:
            x1 = np.arange(len(df_burn))
            x1_smooth = np.linspace(x1.min(), x1.max(), 300)
            plan_spline = make_interp_spline(x1, df_burn['Planned'], k=3)
            actual_spline = make_interp_spline(x1, df_burn['Actual'], k=3)
            plan_smooth = plan_spline(x1_smooth)
            actual_smooth = actual_spline(x1_smooth)

            fig1 = go.Figure()
            fig1.add_trace(go.Scatter(x=x1_smooth, y=plan_smooth, mode='lines', name='Planned',
                                      line=dict(color='deepskyblue', width=2)))
            fig1.add_trace(go.Scatter(x=x1_smooth, y=actual_smooth, mode='lines', name='Actual',
                                      line=dict(color='orange', width=2)))
            fig1.add_annotation(x=x1_smooth[0], y=plan_smooth[0],
                                text=f"Start: {df_burn['Planned'].iloc[0]}",
                                showarrow=True, arrowhead=2, ax=0, ay=-30,
                                font=dict(color='deepskyblue'))
            fig1.add_annotation(x=x1_smooth[-1], y=actual_smooth[-1],
                                text=f"End: {df_burn['Actual'].iloc[-1]}",
                                showarrow=True, arrowhead=2, ax=0, ay=30,
                                font=dict(color='orange'))
            fig1.update_layout(title="E2E - Passed Burndown",
                               xaxis_title="Timeline",
                               yaxis_title="Work Remaining",
                               template="plotly_dark",
                               hovermode="x unified",
                               height=400)
            chart1 = fig1

    # Chart 2: Metric Comparison (E to G)
    chart2 = None
    if df3.shape[1] >= 7:
        df_metric = df3.iloc[:, 4:7].copy()
        df_metric.columns = ['Date2', 'Metric1', 'Metric2']
        df_metric['Date2'] = pd.to_datetime(df_metric['Date2'], errors='coerce')
        df_metric['Metric1'] = pd.to_numeric(df_metric['Metric1'], errors='coerce')
        df_metric['Metric2'] = pd.to_numeric(df_metric['Metric2'], errors='coerce')
        df_metric = df_metric.replace([np.inf, -np.inf], np.nan).dropna()

        if len(df_metric) >= 4:
            x2 = np.arange(len(df_metric))
            x2_smooth = np.linspace(x2.min(), x2.max(), 300)
            metric1_spline = make_interp_spline(x2, df_metric['Metric1'], k=3)
            metric2_spline = make_interp_spline(x2, df_metric['Metric2'], k=3)
            metric1_smooth = metric1_spline(x2_smooth)
            metric2_smooth = metric2_spline(x2_smooth)

            fig2 = go.Figure()
            fig2.add_trace(go.Scatter(x=x2_smooth, y=metric1_smooth, mode='lines', name='Metric1',
                                      line=dict(color='limegreen', width=2)))
            fig2.add_trace(go.Scatter(x=x2_smooth, y=metric2_smooth, mode='lines', name='Metric2',
                                      line=dict(color='magenta', width=2)))
            fig2.add_annotation(x=x2_smooth[0], y=metric1_smooth[0],
                                text=f"Start: {df_metric['Metric1'].iloc[0]}",
                                showarrow=True, arrowhead=2, ax=0, ay=-30,
                                font=dict(color='limegreen'))
            fig2.add_annotation(x=x2_smooth[-1], y=metric2_smooth[-1],
                                text=f"End: {df_metric['Metric2'].iloc[-1]}",
                                showarrow=True, arrowhead=2, ax=0, ay=30,
                                font=dict(color='magenta'))
            fig2.update_layout(title="NFT - Passed Burndown",
                               xaxis_title="Timeline",
                               yaxis_title="Values",
                               template="plotly_dark",
                               hovermode="x unified",
                               height=400)
            chart2 = fig2

    # Display both charts side by side
    if chart1 or chart2:
        col1, col2 = st.columns(2)
        with col1:
            if chart1:
                st.plotly_chart(chart1, use_container_width=True)
        with col2:
            if chart2:
                st.plotly_chart(chart2, use_container_width=True)
else:
    st.error("Excel file not found.")


