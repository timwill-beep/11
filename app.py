import streamlit as st
import pandas as pd
import io
import plotly.express as px

def init_session_state():
    """Initializes the session state with mock data if it doesn't exist."""
    if 'demands' not in st.session_state:
        st.session_state.demands = pd.DataFrame({
            'Part Number': ['P-001', 'P-002', 'P-003'],
            'Name': ['Widget A', 'Widget B', 'Widget C'],
            'AMU (Units/Month)': [500, 1200, 300],
            'Customer': ['Acme Corp', 'Globex', 'Acme Corp'],
            'Product Family': ['Widgets', 'Widgets', 'Gizmos'],
            'Product Sub-fam': ['Standard', 'Standard', 'Premium'],
            'Product': ['Standard Widget', 'Standard Widget', 'Premium Gizmo'],
            'NPD': ['No', 'No', 'Yes'],
            'Priority': ['High', 'Normal', 'High'],
            'Target Launch Date': ['N/A', 'N/A', '2024-Q3']
        })

    if 'routings' not in st.session_state:
        st.session_state.routings = pd.DataFrame({
            'Part Number': ['P-001', 'P-001', 'P-002', 'P-002', 'P-003', 'P-003'],
            'Operation Step': [10, 20, 10, 20, 10, 20],
            'Target Resource': ['Milling_Bucket_1', 'Swiss_Bucket_1', 'Milling_Bucket_1', 'Lathe_01', 'Swiss_Bucket_1', 'Assembly_01'],
            'Time per Unit (Hours)': [0.1, 0.05, 0.08, 0.15, 0.2, 0.5]
        })

    if 'machines' not in st.session_state:
        st.session_state.machines = pd.DataFrame({
            'Machine ID': ['Mill_01', 'Mill_02', 'Swiss_01', 'Swiss_02', 'Swiss_03', 'Lathe_01', 'Assembly_01'],
            'Name': ['Milling Machine 1', 'Milling Machine 2', 'Swiss Lathe 1', 'Swiss Lathe 2', 'Swiss Lathe 3', 'Standard Lathe 1', 'Assembly Station 1'],
            'Bucket ID': ['Milling_Bucket_1', 'Milling_Bucket_1', 'Swiss_Bucket_1', 'Swiss_Bucket_1', 'Swiss_Bucket_1', 'Lathe_Bucket_1', 'Assembly_Bucket_1'],
            'Available Capacity (Hours/Month)': [160, 160, 160, 160, 160, 160, 320]
        })

def generate_excel_template():
    """Generates an Excel template with the correct sheets and headers."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame(columns=[
            'Part Number', 'Name', 'AMU (Units/Month)', 'Customer', 
            'Product Family', 'Product Sub-fam', 'Product', 'NPD', 
            'Priority', 'Target Launch Date'
        ]).to_excel(writer, sheet_name='Demands', index=False)
        
        pd.DataFrame(columns=[
            'Part Number', 'Operation Step', 'Target Resource', 'Time per Unit (Hours)'
        ]).to_excel(writer, sheet_name='Routings', index=False)
        
        pd.DataFrame(columns=[
            'Machine ID', 'Name', 'Bucket ID', 'Available Capacity (Hours/Month)'
        ]).to_excel(writer, sheet_name='Machines', index=False)
    
    return output.getvalue()

def handle_excel_upload(uploaded_file):
    """Processes the uploaded Excel file and updates session state."""
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        required_sheets = ['Demands', 'Routings', 'Machines']
        missing_sheets = [sheet for sheet in required_sheets if sheet not in xls.sheet_names]
        
        if missing_sheets:
            st.error(f"Uploaded file is missing required sheets: {', '.join(missing_sheets)}")
            return False
        
        st.session_state.demands = pd.read_excel(xls, sheet_name='Demands')
        st.session_state.routings = pd.read_excel(xls, sheet_name='Routings')
        st.session_state.machines = pd.read_excel(xls, sheet_name='Machines')
        
        st.success("Data successfully imported from Excel!")
        return True
        
    except Exception as e:
        st.error(f"Error processing the Excel file: {str(e)}")
        return False

def calculate_capacity(demands, routings, machines, baseline_method, manual_baseline_df=None):
    """Calculates required capacity and utilization."""
    
    # 1. Calculate Required Hours from Demands and Routings
    merged_df = pd.merge(demands, routings, on='Part Number', how='inner')
    merged_df['Required Hours'] = merged_df['AMU (Units/Month)'] * merged_df['Time per Unit (Hours)']
    required_hours_by_resource = merged_df.groupby('Target Resource')['Required Hours'].sum().reset_index()
    required_hours_by_resource.rename(columns={'Target Resource': 'Resource'}, inplace=True)
    
    # 2. Add Baseline Load
    total_required = pd.DataFrame(columns=['Resource', 'Total Required Hours'])
    total_required = pd.merge(total_required, required_hours_by_resource, on='Resource', how='outer').fillna(0)
    total_required['Total Required Hours'] = total_required['Total Required Hours'] + total_required['Required Hours']
    total_required.drop(columns=['Required Hours'], inplace=True)

    if baseline_method == "Method A: Manual Entry" and manual_baseline_df is not None:
        manual_baseline_df.rename(columns={'Current Load (Hours/Month)': 'Baseline Hours'}, inplace=True)
        total_required = pd.merge(total_required, manual_baseline_df, on='Resource', how='outer').fillna(0)
        total_required['Total Required Hours'] += total_required['Baseline Hours']
        total_required.drop(columns=['Baseline Hours'], inplace=True)

    # 3. Build Resource Capacity mapping (Buckets and Machines)
    resource_capacity = []
    
    for index, row in machines.iterrows():
        resource_capacity.append({
            'Resource': row['Machine ID'],
            'Type': 'Machine',
            'Parent Bucket': row['Bucket ID'],
            'Available Capacity': row['Available Capacity (Hours/Month)']
        })
        
    bucket_capacities = machines.groupby('Bucket ID')['Available Capacity (Hours/Month)'].sum().reset_index()
    for index, row in bucket_capacities.iterrows():
        resource_capacity.append({
            'Resource': row['Bucket ID'],
            'Type': 'Bucket',
            'Parent Bucket': None,
            'Available Capacity': row['Available Capacity (Hours/Month)']
        })
        
    capacity_df = pd.DataFrame(resource_capacity)

    # 4. Final Aggregation and Utilization Calculation
    results_df = pd.merge(capacity_df, total_required, on='Resource', how='left').fillna(0)
    
    final_results = []
    
    for index, row in results_df.iterrows():
        res_name = row['Resource']
        res_type = row['Type']
        avail_cap = row['Available Capacity']
        
        if res_type == 'Bucket':
            direct_bucket_load = row['Total Required Hours']
            machines_in_bucket = results_df[(results_df['Parent Bucket'] == res_name) & (results_df['Type'] == 'Machine')]
            machine_load_in_bucket = machines_in_bucket['Total Required Hours'].sum()
            
            total_load = direct_bucket_load + machine_load_in_bucket
            
            final_results.append({
                'Resource': res_name,
                'Type': res_type,
                'Available Capacity': avail_cap,
                'Total Required Hours': total_load,
                'Utilization %': (total_load / avail_cap * 100) if avail_cap > 0 else 0
            })
            
        elif res_type == 'Machine':
            total_load = row['Total Required Hours']
            
            final_results.append({
                'Resource': res_name,
                'Type': res_type,
                'Available Capacity': avail_cap,
                'Total Required Hours': total_load,
                'Utilization %': (total_load / avail_cap * 100) if avail_cap > 0 else 0
            })
            
    return pd.DataFrame(final_results)

def render_dashboard(results_df):
    """Renders the bottleneck analysis dashboard."""
    st.markdown("---")
    st.write("### Bottleneck Analysis Dashboard")
    
    if results_df.empty:
        st.warning("No data to analyze. Please ensure you have added Demands, Routings, and Machines.")
        return

    bucket_results = results_df[results_df['Type'] == 'Bucket'].copy()
    machine_results = results_df[results_df['Type'] == 'Machine'].copy()

    col1, col2 = st.columns(2)

    with col1:
        st.write("#### Group (Bucket) Utilization")
        if not bucket_results.empty:
            bucket_results['Status'] = bucket_results['Utilization %'].apply(lambda x: 'Bottleneck (>100%)' if x > 100 else 'Healthy (<=100%)')
            color_map = {'Bottleneck (>100%)': 'red', 'Healthy (<=100%)': 'green'}
            
            fig_bucket = px.bar(
                bucket_results, 
                x='Resource', 
                y='Utilization %', 
                color='Status',
                color_discrete_map=color_map,
                text='Utilization %',
                title="Total Capacity Utilization per Bucket"
            )
            fig_bucket.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
            fig_bucket.add_hline(y=100, line_dash="dash", line_color="black", annotation_text="100% Capacity")
            st.plotly_chart(fig_bucket, use_container_width=True)
        else:
            st.info("No buckets defined.")

    with col2:
        st.write("#### Specific Machine Utilization")
        st.caption("Note: This only shows load routed directly to the specific machine, not shared bucket load.")
        if not machine_results.empty:
            machine_results['Status'] = machine_results['Utilization %'].apply(lambda x: 'Bottleneck (>100%)' if x > 100 else 'Healthy (<=100%)')
            color_map = {'Bottleneck (>100%)': 'red', 'Healthy (<=100%)': 'green'}
            
            fig_machine = px.bar(
                machine_results, 
                x='Resource', 
                y='Utilization %', 
                color='Status',
                color_discrete_map=color_map,
                text='Utilization %',
                title="Utilization per Individual Machine"
            )
            fig_machine.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
            fig_machine.add_hline(y=100, line_dash="dash", line_color="black", annotation_text="100% Capacity")
            st.plotly_chart(fig_machine, use_container_width=True)
        else:
            st.info("No machines defined.")
            
    st.write("#### Detailed Results Table")
    display_df = results_df.copy()
    display_df['Utilization %'] = display_df['Utilization %'].apply(lambda x: f"{x:.1f}%")
    display_df['Total Required Hours'] = display_df['Total Required Hours'].apply(lambda x: f"{x:.1f}")
    display_df['Available Capacity'] = display_df['Available Capacity'].apply(lambda x: f"{x:.1f}")
    
    def highlight_bottlenecks(row):
        util = float(row['Utilization %'].strip('%'))
        if util > 100:
            return ['background-color: rgba(255, 0, 0, 0.2)'] * len(row)
        return [''] * len(row)
        
    st.dataframe(display_df.style.apply(highlight_bottlenecks, axis=1), use_container_width=True)

def main():
    st.set_page_config(page_title="Factory Capacity Planning", layout="wide")
    st.title("Factory Capacity Planning Simulation")
    
    init_session_state()

    st.sidebar.header("Data Management")
    
    template_bytes = generate_excel_template()
    st.sidebar.download_button(
        label="Download Excel Template",
        data=template_bytes,
        file_name="capacity_planning_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.sidebar.markdown("---")
    
    uploaded_file = st.sidebar.file_uploader("Upload Data (Excel)", type=['xlsx'])
    if uploaded_file is not None:
        if st.sidebar.button("Import Data"):
            handle_excel_upload(uploaded_file)

    st.write("### Data Management")
    
    tab1, tab2, tab3 = st.tabs(["Demands", "Routings", "Machines/Buckets"])
    
    with tab1:
        st.write("Manage product demands and average monthly usage (AMU).")
        st.session_state.demands = st.data_editor(
            st.session_state.demands,
            num_rows="dynamic",
            use_container_width=True,
            key="demands_editor"
        )
        
    with tab2:
        st.write("Manage manufacturing steps and target resources (Specific Machine or Bucket).")
        st.session_state.routings = st.data_editor(
            st.session_state.routings,
            num_rows="dynamic",
            use_container_width=True,
            key="routings_editor"
        )
        
    with tab3:
        st.write("Manage machines, assign them to buckets, and set available capacity.")
        st.session_state.machines = st.data_editor(
            st.session_state.machines,
            num_rows="dynamic",
            use_container_width=True,
            key="machines_editor"
        )

    st.markdown("---")
    st.write("### Baseline Load Configuration")
    
    baseline_method = st.radio(
        "Select how you want to configure the current baseline load:",
        ("Method A: Manual Entry", "Method B: Calculated from Current Products"),
        horizontal=True
    )
    
    manual_baseline_df = None
    if baseline_method == "Method A: Manual Entry":
        st.write("Enter the existing load (in Hours/Month) for each machine or bucket.")
        if 'baseline_load' not in st.session_state:
            machines = st.session_state.machines['Machine ID'].tolist()
            buckets = st.session_state.machines['Bucket ID'].unique().tolist()
            resources = list(set(machines + buckets))
            st.session_state.baseline_load = pd.DataFrame({
                'Resource': resources,
                'Current Load (Hours/Month)': [0.0] * len(resources)
            })
            
        manual_baseline_df = st.data_editor(
            st.session_state.baseline_load,
            num_rows="dynamic",
            use_container_width=True,
            key="baseline_editor"
        )
    else:
        st.write("Baseline load will be calculated from the products currently listed in the 'Demands' table.")
        st.info("In this mode, ALL products in the Demands table are considered 'Current'. To simulate NEW demand, add them to the Demands table and observe the changes in the Bottleneck Analysis below.")

    st.markdown("---")
    if st.button("Run Capacity Simulation", type="primary"):
        with st.spinner("Calculating capacity and identifying bottlenecks..."):
            results_df = calculate_capacity(
                st.session_state.demands, 
                st.session_state.routings, 
                st.session_state.machines,
                baseline_method,
                manual_baseline_df
            )
            render_dashboard(results_df)

if __name__ == "__main__":
    main()
