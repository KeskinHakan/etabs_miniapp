import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from etabs_backend import ETABSManager
import plotly.graph_objects as go

st.title("🚀 ETABS Mini Tool: First & Last Station Processor")

# User input for model path
model_path = st.text_input("Enter ETABS Model Path (.edb):", r"C:\Users\hakan\Desktop\ETABS\00_MODEL\SAMPLE.edb")

# --- Initialize session states ---
for key in ['etabs', 'final_df', 'raw_df']:
    if key not in st.session_state:
        st.session_state[key] = None

# --- Connect and Analyze Button ---
if st.button("Connect and Analyze"):
    st.session_state.etabs = ETABSManager(model_path)
    st.session_state.etabs.run_analysis()
    st.success("Model loaded and analysis completed.")

# --- Table Selection ---
allowed_tables = [
    "Element Forces - Beams",
    "Element Forces - Columns",
    "Element Forces - Braces"
]

table_key = st.selectbox("Select a Table to Process", allowed_tables)

# --- Fetch and Process Button ---
if st.button("Fetch and Process Table"):
    if st.session_state.etabs is None:
        st.error("First connect to ETABS!")
    else:
        df = st.session_state.etabs.get_table_as_dataframe(table_key)
        st.session_state.raw_df = df.copy()  # Save full unfiltered version

        # Detect if Beam/Column/Brace
        if "Beam" in table_key:
            id_cols = ["Story", "Beam", "UniqueName", "OutputCase"]
        elif "Column" in table_key:
            id_cols = ["Story", "Column", "UniqueName", "OutputCase"]
        elif "Brace" in table_key:
            id_cols = ["Story", "Brace", "UniqueName", "OutputCase"]
        else:
            id_cols = ["Story", "UniqueName", "OutputCase"]

        st.session_state.final_df = st.session_state.etabs.process_first_last_station(df, id_cols)
        st.success("Table fetched and processed!")

# --- Display Table if available ---
if st.session_state.final_df is not None:
    st.subheader("📋 First & Last Station Table")
    st.dataframe(st.session_state.final_df)

    # CSV Download
    csv = st.session_state.final_df.to_csv(index=False).encode('utf-8')
    st.download_button("Download as CSV", csv, "processed_table.csv", "text/csv")

    # Excel Download (optional)
    excel_file = "processed_table.xlsx"
    with pd.ExcelWriter(excel_file) as writer:
        st.session_state.final_df.to_excel(writer, index=False)
    with open(excel_file, "rb") as file:
        st.download_button("Download as Excel", file, "processed_table.xlsx", "application/vnd.ms-excel")

# --- Live Frame Force Visualization for Selected Members ---
if st.session_state.etabs is not None:
    st.subheader("📈 Visualize Forces from FrameForce Live Data")

    # --- Button to Get FrameForce Data for current selection ---
    if st.button("Get FrameForce Data for Selected Frames"):
        try:
            frameforce_df = st.session_state.etabs.get_frameforce_df()
            st.session_state.frameforce_df = frameforce_df
            st.success("FrameForce data retrieved successfully for selected frames.")
        except Exception as e:
            st.error(f"Failed to get FrameForce data: {e}")

    # --- If we have data, show the visualization options ---
    if st.session_state.get('frameforce_df') is not None:
        frameforce_df = st.session_state.frameforce_df

        # Dynamically fill FrameName
        frame_names = sorted(frameforce_df["FrameName"].unique())
        frame_selected = st.selectbox("Select Frame Name", frame_names)

        # Filter OutputCases only available for this Frame
        output_cases = sorted(frameforce_df[frameforce_df["FrameName"] == frame_selected]["OutputCase"].unique())
        output_case_selected = st.selectbox("Select OutputCase (Combo)", output_cases)

        # Filter StepTypes only available for this Frame + OutputCase
        step_types = sorted(frameforce_df[
            (frameforce_df["FrameName"] == frame_selected) &
            (frameforce_df["OutputCase"] == output_case_selected)
        ]["StepType"].unique())

        step_type_selected = st.selectbox("Select Step Type", step_types)

        force_column = st.selectbox("Select Force Type", ["P", "V2", "V3", "T", "M2", "M3"])

        if st.button("🎯 Visualize Force Distribution"):
            try:
                filtered_df = frameforce_df[
                    (frameforce_df["FrameName"] == frame_selected) &
                    (frameforce_df["OutputCase"] == output_case_selected) &
                    (frameforce_df["StepType"] == step_type_selected)
                ].sort_values("ObjStation")

                fig = go.Figure()

                fig.add_trace(go.Scatter(
                    x=filtered_df["ObjStation"],
                    y=filtered_df[force_column],
                    mode='lines+markers',
                    name=f'{force_column} ({step_type_selected})'
                ))

                # Add zero line
                fig.add_hline(y=0, line_dash="dot", line_color="gray", annotation_text="Zero", annotation_position="top left")

                fig.update_layout(
                    title=f"{force_column} Distribution - Frame {frame_selected} - {output_case_selected} ({step_type_selected})",
                    xaxis_title="Station [m]",
                    yaxis_title=f"{force_column} [kN or kNm]",
                    template="plotly_white",
                    height=500,
                )

                st.plotly_chart(fig, use_container_width=True)

            except Exception as e:
                st.error(f"Plot failed: {e}")




# --- Close ETABS Button ---
if st.button("Close ETABS"):
    if st.session_state.etabs is not None:
        st.session_state.etabs.close_etabs()
        st.success("ETABS Closed.")
        for key in ['etabs', 'final_df', 'raw_df']:
            st.session_state[key] = None
