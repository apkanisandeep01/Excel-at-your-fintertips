import streamlit as st
import pandas as pd

# Dependency checks
try:
    import openpyxl
except ImportError:
    st.error("The 'openpyxl' library is missing. Install it with 'pip install openpyxl'.")
    st.stop()

try:
    import xlrd
except ImportError:
    st.error("The 'xlrd' library is missing. Install it with 'pip install xlrd' for .xls support.")
    st.stop()

# Initialize session state
if 'uploaded_files' not in st.session_state:
    st.session_state.uploaded_files = []
if 'dataframes' not in st.session_state:
    st.session_state.dataframes = []

# Title
st.title("Python-powered Excel Tools")

# Sidebar for controls
with st.sidebar:
    st.header("Controls")
    st.subheader("Precision. Speed. Control. Data, at your fingertips.")
    st.info("Supported file types: .xlsx, .xls, .csv")
    
    # File uploader
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
            st.session_state.clear()
            st.session_state.uploaded_files = []
            st.session_state.dataframes = []
            st.rerun()
    with col2:
        if st.button("Refresh"):
            st.session_state.dataframes = []
            st.rerun()

    # Operation selection
    if st.session_state.uploaded_files:
        option = st.selectbox(
            "Choose an operation",
            ["View Only", "Combine", "Split Excel", "Drop Columns", "Join Tables"],
            key="operation_select"
        )

# Main area
st.header("Your Dynamic View")

# Process uploaded files
if st.session_state.uploaded_files and not st.session_state.dataframes:
    for file in st.session_state.uploaded_files:
        file_name = file.name.lower()
        try:
            if file_name.endswith('xlsx'):
                df = pd.read_excel(file, engine="openpyxl")
                st.session_state.dataframes.append(df)
                st.success(f"Successfully loaded: {file_name}")
            elif file_name.endswith('xls'):
                df = pd.read_excel(file, engine="xlrd")
                st.session_state.dataframes.append(df)
                st.success(f"Successfully loaded: {file_name}")
            elif file_name.endswith('csv'):
                df = pd.read_csv(file)
                st.session_state.dataframes.append(df)
                st.success(f"Successfully loaded: {file_name}")
            else:
                st.error(f"Unsupported file format: {file_name}")
        except Exception as e:
            st.error(f"Error processing {file.name}: {str(e)}")

# Operations
if st.session_state.dataframes:
    if option == "View Only":
        st.subheader("Uploaded Data:")
        for i, df in enumerate(st.session_state.dataframes, 1):
            st.write(f"File {i} ({st.session_state.uploaded_files[i-1].name}):")
            st.write(f"Rows: {df.shape[0]} | Columns: {df.shape[1]} | Duplicates: {df.duplicated().sum()}")
            st.dataframe(df)
            st.write("---")

    elif option == "Combine":
        if len(st.session_state.dataframes) < 2:
            st.warning("Please upload at least 2 files to combine.")
        else:
            try:
                st.subheader("Combined Dataframe:")
                combined_df = pd.concat(st.session_state.dataframes, ignore_index=True)
                st.write(f"Files Combined: {len(st.session_state.dataframes)}")
                st.write(f"Rows: {combined_df.shape[0]} | Columns: {combined_df.shape[1]}")
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
        st.subheader("Split Excel")
        if len(st.session_state.dataframes) > 1:
            file_to_split = st.selectbox(
                "Select file to split",
                options=[f"File {i+1} ({st.session_state.uploaded_files[i].name})" for i in range(len(st.session_state.uploaded_files))],
                key="split_file_select"
            )
            df_index = int(file_to_split.split()[1]) - 1
        else:
            df_index = 0
        df_to_split = st.session_state.dataframes[df_index]

        column_to_split = st.selectbox(
            "Select column to split by",
            options=df_to_split.columns.tolist(),
            key="split_column_select"
        )

        if column_to_split:
            st.write(f"Found {len(df_to_split[column_to_split].unique())} unique values in '{column_to_split}'")
            for value in df_to_split[column_to_split].unique():
                if pd.notna(value):
                    split_df = df_to_split[df_to_split[column_to_split] == value]
                    st.write(f"Data for {column_to_split} = {value} (Rows: {split_df.shape[0]}):")
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
        st.subheader("Drop Columns")
        if len(st.session_state.dataframes) > 1:
            file_to_modify = st.selectbox(
                "Select file to modify",
                options=[f"File {i+1} ({st.session_state.uploaded_files[i].name})" for i in range(len(st.session_state.uploaded_files))],
                key="drop_file_select"
            )
            df_index = int(file_to_modify.split()[1]) - 1
        else:
            df_index = 0
        df_to_modify = st.session_state.dataframes[df_index].copy()

        columns_to_drop = st.multiselect(
            "Select columns to drop",
            options=df_to_modify.columns.tolist(),
            default=[],
            key="drop_columns_select"
        )

        if columns_to_drop:
            try:
                modified_df = df_to_modify.drop(columns=columns_to_drop)
                st.write(f"Rows: {modified_df.shape[0]} | Columns: {modified_df.shape[1]}")
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
        else:
            st.write("Select columns to drop to see the result.")

    elif option == "Join Tables":
        st.subheader("Join Tables")
        if len(st.session_state.dataframes) != 2:
            st.warning("Please upload exactly 2 files to perform a join.")
        else:
            df1 = st.session_state.dataframes[0]
            df2 = st.session_state.dataframes[1]
            left_on = st.selectbox(
                "Column from File 1 (left table)",
                options=df1.columns.tolist(),
                key="left_on_select"
            )
            right_on = st.selectbox(
                "Column from File 2 (right table)",
                options=df2.columns.tolist(),
                key="right_on_select"
            )

            if left_on and right_on:
                try:
                    merged_df = df1.merge(df2, how="inner", left_on=left_on, right_on=right_on)
                    if merged_df.empty:
                        st.warning("Join resulted in an empty table. Check column compatibility.")
                    else:
                        st.write(f"Rows: {merged_df.shape[0]} | Columns: {merged_df.shape[1]}")
                        st.dataframe(merged_df)
                        csv = merged_df.to_csv(index=False)
                        st.download_button(
                            label="Download Joined CSV",
                            data=csv,
                            file_name="joined_data.csv",
                            mime="text/csv"
                        )
                except ValueError as e:
                    st.error(f"Join failed. Check column data types: {str(e)}")
                except Exception as e:
                    st.error(f"Error joining tables: {str(e)}")

# Footer
name = "Apkani Sandeep Kumar"
email = "apkanisandeep00@outlook.com"
footer = f"""
    <footer style='text-align: left; padding: 10px; width: 100%; background-color: #000000; color: #ffffff; font-weight: bold;'>
        <p>Created by {name} | Email: {email} | <a href='https://www.datascienceportfol.io/apkanisandeep' style='color: #ffffff; text-decoration: none;'>Visit my page 🌐</a> | © 2025</p>
    </footer>
"""
st.markdown(footer, unsafe_allow_html=True)
