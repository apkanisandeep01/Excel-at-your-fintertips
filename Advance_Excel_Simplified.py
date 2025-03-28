import streamlit as st
import pandas as pd

# Check for openpyxl dependency
try:
    import openpyxl
    from packaging import version
    required_version = "3.1.0"
    installed_version = openpyxl.__version__
    if version.parse(installed_version) < version.parse(required_version):
        st.error(f"Pandas requires openpyxl version {required_version} or newer. You have {installed_version}. Please upgrade using 'pip install --upgrade openpyxl'.")
        st.stop()
except ImportError:
    st.error("The 'openpyxl' library is missing. Please install it using 'pip install openpyxl' and restart the app.")
    st.stop()
    
# Initialize session state
if 'uploaded_files' not in st.session_state:
    st.session_state.uploaded_files = []
if 'dataframes' not in st.session_state:
    st.session_state.dataframes = []

st.title("Python-powered Excel tools")

# Sidebar for controls
with st.sidebar:
    st.header("Controls")
    st.subheader('Precision. Speed. Control. Data, at your fingertips. Work smarter, faster, better.')
    st.info("Supported file types: .xlsx, .xls, .csv")
    # Get the number of files
    count = st.number_input("How many files do you want to upload?", min_value=1, step=1, key="file_count")

    # File upload section
    uploaded_files = st.file_uploader(
        "Upload your files",
        type=['xlsx', 'xls', 'csv'],
        accept_multiple_files=True,
        key="file_uploader"
    )
    
    # Update session state with uploaded files
    if uploaded_files:
        st.session_state.uploaded_files = uploaded_files

    # Restart and Refresh buttons
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Restart"):
            # Clear all session state variables
            st.session_state.clear()
            # Reinitialize essential variables
            st.session_state.uploaded_files = []
            st.session_state.dataframes = []
            # Reset file_count by removing its key (will revert to default on rerun)
            if "file_count" in st.session_state:
                del st.session_state["file_count"]
            # Rerun the app to reset UI
            st.rerun()
    with col2:
        if st.button("Refresh"):
            st.session_state.dataframes = []  # Clear processed dataframes only
            st.rerun()
   
    # Processing options
    if st.session_state.uploaded_files:
        option = st.selectbox("Choose an operation", ["View Only", "Combine", "Split Excel", "Drop Columns", "Join Tables"], key="operation_select")

# Main area for results
st.header("Your Dynamic View")
# Process uploaded files
if st.session_state.uploaded_files:
    if len(st.session_state.uploaded_files) != count:
        st.warning(f"You specified {count} files but uploaded {len(st.session_state.uploaded_files)}. Processing all uploaded files anyway.")
    
    # Only process files if dataframes are empty
    if not st.session_state.dataframes:
        for file in st.session_state.uploaded_files:
            file_name = file.name.lower()
            try:
                if file_name.endswith(('xlsx', 'xls')):
                    df = pd.read_excel(file)
                    st.session_state.dataframes.append(df)
                    st.success(f"Successfully loaded Excel file: {file_name}")
                elif file_name.endswith('csv'):
                    df = pd.read_csv(file)
                    st.session_state.dataframes.append(df)
                    st.success(f"Successfully loaded CSV file: {file_name}")
                else:
                    st.error(f"Unsupported file format: {file_name}")
            except Exception as e:
                st.error(f"Error processing {file.name}: {str(e)}")

    # Display basic info if "View Only"
    if st.session_state.dataframes and option == "View Only":
        st.subheader("Uploaded Data:")
        for i, df in enumerate(st.session_state.dataframes, 1):
            st.write(f"File {i} ({st.session_state.uploaded_files[i-1].name}):")
            st.write(f"Number of Rows: {df.shape[0]}")
            st.write(f"Number of Columns: {df.shape[1]}")
            st.write(f"Number of Duplicate Rows: {df.duplicated().sum()}")
            st.dataframe(df)
            st.write("---")

    # Process operations
    if option == "Combine":
        try:
            if len(st.session_state.dataframes) < 2:
                st.warning("Please upload at least 2 files to combine.")
            else:
                st.write(f"Combining {len(st.session_state.dataframes)} files (unlimited files supported)...")
                combined_df = pd.concat(st.session_state.dataframes, ignore_index=True)
                st.subheader("Combined Dataframe:")
                st.write(f"Total Files Combined: {len(st.session_state.dataframes)} (supports unlimited files)")
                st.write(f"Total Rows: {combined_df.shape[0]}")
                st.write(f"Total Columns: {combined_df.shape[1]}")
                st.dataframe(combined_df)
                csv = combined_df.to_csv(index=False)
                st.download_button(
                    label="Download Combined CSV",
                    data=csv,
                    file_name="combined_data.csv",
                    mime="text/csv"
                )
        except Exception as e:
            st.error(f"Error combining files: {str(e)}")

    elif option == "Split Excel":
        with st.sidebar:
            if len(st.session_state.dataframes) > 1:
                file_to_split = st.selectbox(
                    "Select file to split",
                    options=[f"File {i+1} ({st.session_state.uploaded_files[i].name})" for i in range(len(st.session_state.uploaded_files))],
                    key="split_file_select"
                )
                df_index = int(file_to_split.split()[1]) - 1
                df_to_split = st.session_state.dataframes[df_index]
            else:
                df_index = 0
                df_to_split = st.session_state.dataframes[0]

            column_to_split = st.selectbox(
                "Select column to split by",
                options=df_to_split.columns.tolist(),
                key="split_column_select"
            )

        if column_to_split:
            st.subheader(f"Split Results (by {column_to_split}):")
            unique_values = df_to_split[column_to_split].unique()
            st.write(f"Found {len(unique_values)} unique values")
            for value in unique_values:
                if pd.notna(value):
                    split_df = df_to_split[df_to_split[column_to_split] == value]
                    st.write(f"Data for {column_to_split} = {value}:")
                    st.write(f"Rows: {split_df.shape[0]}")
                    st.dataframe(split_df)
                    csv = split_df.to_csv(index=False)
                    st.download_button(
                        label=f"Download {value} CSV",
                        data=csv,
                        file_name=f"split_{column_to_split}_{value}.csv",
                        mime="text/csv"
                    )
                st.write("---")

    elif option == "Drop Columns":
        with st.sidebar:
            if len(st.session_state.dataframes) > 1:
                file_to_modify = st.selectbox(
                    "Select file to drop columns from",
                    options=[f"File {i+1} ({st.session_state.uploaded_files[i].name})" for i in range(len(st.session_state.uploaded_files))],
                    key="drop_file_select"
                )
                df_index = int(file_to_modify.split()[1]) - 1
                df_to_modify = st.session_state.dataframes[df_index].copy()
            else:
                df_index = 0
                df_to_modify = st.session_state.dataframes[0].copy()

            columns_to_drop = st.multiselect(
                "Select columns to drop",
                options=df_to_modify.columns.tolist(),
                default=[],
                key="drop_columns_select"
            )

        if columns_to_drop:
            try:
                st.subheader("Drop Columns Result:")
                modified_df = df_to_modify.drop(columns=columns_to_drop, inplace=False)
                st.write("Modified Dataframe after dropping columns:")
                st.write(f"Rows: {modified_df.shape[0]}")
                st.write(f"Columns: {modified_df.shape[1]}")
                st.dataframe(modified_df)
                
                csv = modified_df.to_csv(index=False)
                st.download_button(
                    label="Download Modified CSV",
                    data=csv,
                    file_name=f"modified_{st.session_state.uploaded_files[df_index].name.split('.')[0]}.csv",
                    mime="text/csv"
                )
            except Exception as e:
                st.error(f"Error dropping columns: {str(e)}")
        elif columns_to_drop == [] and option == "Drop Columns":
            st.write("Select columns from the sidebar to see the modified dataframe")

    elif option == "Join Tables":
        with st.sidebar:
            if len(st.session_state.dataframes) != 2:
                st.warning("Please upload exactly 2 files to perform a join.")
            else:
                df1 = st.session_state.dataframes[0]
                df2 = st.session_state.dataframes[1]
                left_on = st.selectbox(
                    "Select column from File 1 (left table)",
                    options=df1.columns.tolist(),
                    key="left_on_select"
                )
                right_on = st.selectbox(
                    "Select column from File 2 (right table)",
                    options=df2.columns.tolist(),
                    key="right_on_select"
                )

        if len(st.session_state.dataframes) == 2 and left_on and right_on:
            try:
                st.subheader("Joined Dataframe:")
                df1 = st.session_state.dataframes[0]
                df2 = st.session_state.dataframes[1]
                merged_df = df1.merge(df2, how="inner", left_on=left_on, right_on=right_on)
                st.write(f"Total Rows: {merged_df.shape[0]}")
                st.write(f"Total Columns: {merged_df.shape[1]}")
                st.dataframe(merged_df)
                
                csv = merged_df.to_csv(index=False)
                st.download_button(
                    label="Download Joined CSV",
                    data=csv,
                    file_name="joined_data.csv",
                    mime="text/csv"
                )
            except Exception as e:
                st.error(f"Error joining tables: {str(e)}")

# Footer
name = "Apkani Sandeep Kumar"
email = "apkanisandeep00@outlook.com"
footer = f"""
    <footer style='text-align: left; padding: 10px; position: fixed; bottom: 0; width: 100%; background-color: #000000; color: #ffffff; font-weight: bold;'>
        <p>Created by {name} | Email: {email} | <a href='https://www.datascienceportfol.io/apkanisandeep' style='color: #ffffff; text-decoration: none;'>Visit my page 🌐</a> | © 2025</p>
    </footer>
"""
st.markdown(footer, unsafe_allow_html=True)
