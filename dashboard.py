import streamlit as st
import pandas as pd
import plotly.graph_objs as go
import os

st.set_page_config(page_title="DTR Outage KPIs Dashboard", layout="wide")

# === DTR info: sheet mapping as per your structure ===
dtr_info = {
    "7088-57": {
        "master_file": "7088_32_57_86.xlsx",
        "master_sheet": "Sheet1",
        "outage_file": "7088-57.xlsx",
        "outage_sheet": "O_154",
        "untagged_sheet": "N_T_21",
        "wrongly_mapped_sheet": "W_M_4",
        "feeder": "7088",
        "dtr": "57"
    },
    "7088-32": {
        "master_file": "7088_32_57_86.xlsx",
        "master_sheet": "Sheet1",
        "outage_file": "7088-32.xlsx",
        "outage_sheet": "O_37",
        "untagged_sheet": "N_T_5",
        "wrongly_mapped_sheet": "W_M_32",
        "feeder": "7088",
        "dtr": "32"
    },
    "7088-86": {
        "master_file": "7088_32_57_86.xlsx",
        "master_sheet": "Sheet1",
        "outage_file": "7088-86.xlsx",
        "outage_sheet": "O_165",
        "untagged_sheet": "N_T_33",
        "wrongly_mapped_sheet": "W_M_26",
        "feeder": "7088",
        "dtr": "86"
    },
    "15631-34": {
        "master_file": "15631_34.xlsx",
        "master_sheet": "Sheet1",
        "outage_file": "15631-34.xlsx",
        "outage_sheet": "O_347",
        "untagged_sheet": "N_T_65",
        "wrongly_mapped_sheet": "W_M_40",
        "feeder": "15631",
        "dtr": "34"
    }
}

# Consumption files mapping (file names only, all in current directory)
consumption_files = {
    "7088-57": "7088-57-consumption.xlsx",
    "7088-32": "7088-32 consumption.xlsx",
    "7088-86": "7088-86 consumption.xlsx",
    "15631-34": "15631-34 consumption.xlsx"
}

# --- For dropdowns ---
feeder_options = sorted(set(x['feeder'] for x in dtr_info.values()))
feeder_to_dtr = {}
for dtr_key, v in dtr_info.items():
    feeder_to_dtr.setdefault(v['feeder'], []).append(v['dtr'])

# --- SIDEBAR FOR SELECTION ---
st.sidebar.title("🔌 Select Feeder & DTR")
selected_feeder = st.sidebar.selectbox("Feeder", feeder_options)
selected_dtr = st.sidebar.selectbox("DTR", sorted(feeder_to_dtr[selected_feeder]))

# --- Lookup key for dtr_info ---
dtr_selection = f"{selected_feeder}-{selected_dtr}"
d = dtr_info[dtr_selection]

# --- LOAD DATA ---
try:
    master_all = pd.read_excel(d['master_file'], sheet_name=d['master_sheet'])
    master = master_all[(master_all['dtrcode'] == int(d['dtr'])) & (master_all['Feedercode'] == int(d['feeder']))]
except Exception as e:
    st.error(f"Error loading or filtering master: {e}")
    st.stop()
try:
    outage = pd.read_excel(d['outage_file'], sheet_name=d['outage_sheet'])
except Exception as e:
    st.error(f"Error loading outage sheet: {e}")
    st.stop()
try:
    untagged = pd.read_excel(d['outage_file'], sheet_name=d['untagged_sheet'])
except Exception as e:
    st.error(f"Error loading untagged sheet: {e}")
    st.stop()
try:
    wrongly_mapped = pd.read_excel(d['outage_file'], sheet_name=d['wrongly_mapped_sheet'])
except Exception as e:
    st.error(f"Error loading wrongly mapped sheet: {e}")
    st.stop()

# --- Calculate KPIs ---
kpi1_master_tagged = len(master)
kpi2_connected_outage = len(outage)
kpi3_untagged = len(untagged)
kpi4_wrongly_mapped = len(wrongly_mapped)
kpi5_total_corrected = kpi2_connected_outage + kpi4_wrongly_mapped

# --- Dashboard layout ---
st.markdown(f"""
    <h1 style='color:#1e3799;font-weight:700;margin-bottom:6px'>
        ⚡ DTR Outage KPIs Dashboard <span style='font-size:18px;'>[Feeder: {selected_feeder}, DTR: {selected_dtr}]</span>
    </h1>
    <div style='color:#555;font-size:18px;margin-bottom:24px'>
        Consumer mapping, outage verification, and correction dashboard for <b>DTR {selected_feeder}-{selected_dtr}</b>.
    </div>
""", unsafe_allow_html=True)

# --- KPI Cards ---
st.markdown("### 🏆 Core KPIs at a Glance")
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("📒 Master Tagged Consumers", kpi1_master_tagged)
col2.metric("🟢 Connected (Outage File)", kpi2_connected_outage)
col3.metric("🚫 Untagged (Master Only)", kpi3_untagged)
col4.metric("🔄 Wrongly Mapped (Other DTR, Same Feeder)", kpi4_wrongly_mapped)
col5.metric("🏆 Total After Correction", kpi5_total_corrected)

# --- Bar Chart ---
kpi_labels = [
    "Master Tagged",
    "Connected (Outage)",
    "Untagged",
    "Wrongly Mapped",
    "Total Corrected"
]
kpi_values = [
    kpi1_master_tagged,
    kpi2_connected_outage,
    kpi3_untagged,
    kpi4_wrongly_mapped,
    kpi5_total_corrected
]
bar_colors = ['#0984e3', '#27ae60', '#e74c3c', '#f39c12', '#9b59b6']
fig = go.Figure(data=[
    go.Bar(
        x=kpi_labels,
        y=kpi_values,
        marker_color=bar_colors,
        text=kpi_values,
        textposition='auto'
    )
])
fig.update_layout(
    title=f"DTR Outage KPIs Breakdown ({selected_feeder}-{selected_dtr})",
    yaxis_title="Number of Consumers",
    xaxis_title="KPI Category",
    bargap=0.3
)
st.plotly_chart(fig, use_container_width=True)

# --- Details download ---
st.markdown("### 🗂️ Downloadable Detailed Lists")

with st.expander("Master Tagged Consumers (Sheet1, filtered for selected DTR)"):
    st.dataframe(master, use_container_width=True)
    st.download_button(
        "Download as CSV",
        data=master.to_csv(index=False),
        file_name=f"{selected_feeder}-{selected_dtr}_master_tagged_consumers.csv",
        mime="text/csv"
    )

with st.expander("Connected (Outage File)"):
    st.dataframe(outage, use_container_width=True)
    st.download_button(
        "Download as CSV",
        data=outage.to_csv(index=False),
        file_name=f"{selected_feeder}-{selected_dtr}_connected_outage.csv",
        mime="text/csv"
    )

with st.expander("Untagged (Master Only)"):
    st.dataframe(untagged, use_container_width=True)
    st.download_button(
        "Download as CSV",
        data=untagged.to_csv(index=False),
        file_name=f"{selected_feeder}-{selected_dtr}_untagged_master.csv",
        mime="text/csv"
    )

with st.expander("Wrongly Mapped (Other DTR, Same Feeder)"):
    st.dataframe(wrongly_mapped, use_container_width=True)
    st.download_button(
        "Download as CSV",
        data=wrongly_mapped.to_csv(index=False),
        file_name=f"{selected_feeder}-{selected_dtr}_wrongly_mapped.csv",
        mime="text/csv"
    )
consumption_file = consumption_files.get(dtr_selection)
# --------- CONSUMPTION TREND PLOT (ALWAYS FIRST SHEET) ---------
if consumption_file and os.path.exists(consumption_file):
    df_cons = pd.read_excel(consumption_file, sheet_name=0)

    # Convert date column
    if 'reading_date' in df_cons.columns:
        df_cons['reading_date'] = pd.to_datetime(df_cons['reading_date']).dt.date
    else:
        st.warning("Column 'reading_date' not found in consumption file.")
        st.stop()

    # Chart 1: Meter Count and DLP Loss %
    if 'meter_count' in df_cons.columns and 'DLP_LOSS' in df_cons.columns:
        st.markdown("### 📈 Meter Count and DLP Loss % Trend")
        fig_meter_dlp = go.Figure()
        fig_meter_dlp.add_trace(go.Scatter(
            x=df_cons['reading_date'], y=df_cons['meter_count'],
            mode='lines+markers', name='Meter Count', line=dict(color='green', width=3)
        ))
        fig_meter_dlp.add_trace(go.Scatter(
            x=df_cons['reading_date'], y=df_cons['DLP_LOSS'],
            mode='lines+markers', name='DLP Loss %', line=dict(color='orange', width=3), yaxis='y2'
        ))
        fig_meter_dlp.update_layout(
            title=f"{dtr_selection} - Meter Count and DLP Loss %",
            xaxis_title="Reading Date",
            yaxis=dict(title="Meter Count"),
            yaxis2=dict(
                title="DLP Loss %",
                overlaying="y",
                side="right"
            ),
            legend=dict(x=0.5, y=1.1, orientation='h', xanchor='center')
        )
        st.plotly_chart(fig_meter_dlp, use_container_width=True)

        st.markdown("#### 📋 Meter Count & DLP Loss Table")
        st.dataframe(df_cons[['reading_date', 'meter_count', 'DLP_LOSS']], use_container_width=True)
    else:
        st.warning("Columns 'meter_count' or 'DLP_LOSS' not found in consumption file.")

    # Chart 2: OWM_BLP_Count and BLP Loss %
    if 'OWM_BLP_Count' in df_cons.columns and 'BLP_LOSS' in df_cons.columns:
        st.markdown("### 📉 OWM_BLP_Count and BLP Loss % Trend")
        fig_blp = go.Figure()
        fig_blp.add_trace(go.Scatter(
            x=df_cons['reading_date'], y=df_cons['OWM_BLP_Count'],
            mode='lines+markers', name='OWM_BLP_Count', line=dict(color='blue', width=3)
        ))
        fig_blp.add_trace(go.Scatter(
            x=df_cons['reading_date'], y=df_cons['BLP_LOSS'],
            mode='lines+markers', name='BLP Loss %', line=dict(color='red', width=3), yaxis='y2'
        ))
        fig_blp.update_layout(
            title=f"{dtr_selection} - OWM_BLP_Count and BLP Loss %",
            xaxis_title="Reading Date",
            yaxis=dict(title="OWM_BLP_Count"),
            yaxis2=dict(
                title="BLP Loss %",
                overlaying="y",
                side="right"
            ),
            legend=dict(x=0.5, y=1.1, orientation='h', xanchor='center')
        )
        st.plotly_chart(fig_blp, use_container_width=True)

        st.markdown("#### 📋 OWM_BLP_Count & BLP Loss Table")
        st.dataframe(df_cons[['reading_date', 'OWM_BLP_Count', 'BLP_LOSS']], use_container_width=True)
    else:
        st.warning("Columns 'OWM_BLP_Count' or 'BLP_LOSS' not found in consumption file.")
else:
    st.info(f"No consumption file found for this DTR. (Expected file: {consumption_file if consumption_file else 'N/A'})")

# if consumption_file and os.path.exists(consumption_file):
#     df_cons = pd.read_excel(consumption_file, sheet_name=0)
#     # Smart column name search
#     def smart_col_search(df, search_words):
#         for col in df.columns:
#             name = col.lower().replace(' ', '').replace('_', '')
#             if all(word in name for word in search_words):
#                 return col
#         return None

#     date_col = smart_col_search(df_cons, ['date'])
#     meter_col = smart_col_search(df_cons, ['meter', 'count'])
#     loss_col = smart_col_search(df_cons, ['loss'])

#     if date_col and meter_col and loss_col:
#         df_cons[date_col] = pd.to_datetime(df_cons[date_col]).dt.date

#         st.markdown("### 📈 Meter Count and Loss % Trend (Daily)")
#         fig2 = go.Figure()
#         fig2.add_trace(go.Scatter(
#             x=df_cons[date_col], y=df_cons[meter_col],
#             mode='lines+markers', name='Meter Count', line=dict(color='green', width=3)
#         ))
#         fig2.add_trace(go.Scatter(
#             x=df_cons[date_col], y=df_cons[loss_col],
#             mode='lines+markers', name='%Loss_DLP', line=dict(color='orange', width=3), yaxis='y2'
#         ))

#         fig2.update_layout(
#             xaxis_title="Date",
#             yaxis=dict(title="Meter Count"),
#             yaxis2=dict(
#                 title="%Loss_DLP",
#                 anchor="x",
#                 overlaying="y",
#                 side="right"
#             ),
#             title=f"{dtr_selection} Meter Count and Loss % Trend",
#             legend=dict(x=0.5, y=1.1, orientation='h', xanchor='center')
#         )
#         st.plotly_chart(fig2, use_container_width=True)
#         # Table below chart
#         st.markdown("#### 📋 Daily Meter Count & Loss % Table")
#         table_df = df_cons[[date_col, meter_col, loss_col]].copy()
#         table_df.columns = ['Date', 'Meter Count', '%Loss_DLP']  # Clean labels
#         st.dataframe(table_df, use_container_width=True)
    
#     else:
#         st.info(f"Consumption file found but required columns not detected. Columns found: {df_cons.columns.tolist()}")
# else:
#     st.info("No consumption file found for this DTR. (Expected file: {})".format(consumption_file if consumption_file else "N/A"))

st.markdown("""
    <div style='text-align:center;margin-top:24px;font-size:17px;color:#7f8c8d;'>
        🚀 <b>Power Analytics Dashboard</b> | <i>Esyasoft</i>
    </div>
""", unsafe_allow_html=True)
