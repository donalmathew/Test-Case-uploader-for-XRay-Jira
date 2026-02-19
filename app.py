import streamlit as st
import pandas as pd
import numpy as np
import io


st.title("Excel Automation Tool for making files X-Ray Ready")
st.write("Upload your Excel file to add necessary columns to add multi-Step Test Cases to X-Ray in Jira.")
st.write("""Required columns in the file to be uploaded here: 
1. Summary(Test Case Id with Test Case Description)(This will be the default reference column(see Configuration))
2. Description(Test Objective and Pre-requisites combined with respective headings)
3. Action
4. Test Data
5. Expected Results

Note: Edit the configuration fields as required(copy correct values from your Project in Jira)
""")


with st.sidebar:
    st.header("Configuration")
    ref_col_name = st.text_input("Reference Column Name", "Summary")
    test_type_value = st.text_input("Test Type", "Manual")
    phase_value = st.text_input("Phase", "Testing")
    assignee_id_value = st.text_input("Assignee ID", "your_assignee_id")
    component_name_value = st.text_input("Component Names", "Wealthify")


uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'])

if uploaded_file is not None:
    if st.button("Process File"):
        try:
           
            df = pd.read_excel(uploaded_file)

          
            
           
            df[ref_col_name] = df[ref_col_name].replace(r'^\s*$', np.nan, regex=True)

            ref_col_data = df[ref_col_name]
            test_id_data = df[ref_col_name].ffill()

            ref_col_index = df.columns.get_loc(ref_col_name)

            
            if "Test ID" not in df.columns:
                df.insert(ref_col_index, "Test ID", test_id_data)
                
                ref_col_index = df.columns.get_loc(ref_col_name)

           
            
            new_cols = [
                ("Test Type", test_type_value, 1),
                ("Phase", phase_value, 2),
                ("Assignee ID", assignee_id_value, 3),
                ("Component Names", component_name_value, 4)
            ]

            for col_name, val, offset in new_cols:
                if col_name not in df.columns:
                    
                    data = np.where(ref_col_data.notna(), val, None)
                    
                    df.insert(ref_col_index + offset, col_name, data)

            

            st.success("Success! File processed.")

           
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            
            st.download_button(
                label="Download Processed File",
                data=buffer.getvalue(),
                file_name=f"filled_{uploaded_file.name}",
                mime="application/vnd.ms-excel"
            )

        except KeyError as e:
            st.error(f"Error: Column not found. Check your 'Reference Column Name' setting. ({e})")
        except Exception as e:

            st.error(f"An error occurred: {e}")
