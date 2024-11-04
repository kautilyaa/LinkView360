import streamlit as st
import pandas as pd
import os
from fuzzywuzzy import process, fuzz

EXCEL_FOLDER = 'excel_folder'
if not os.path.exists(EXCEL_FOLDER):
    os.makedirs(EXCEL_FOLDER)

def load_excel_files():
    files = [f for f in os.listdir(EXCEL_FOLDER) if f.endswith('.xlsx')]
    return files

def load_all_sheets(file_path):
    sheet_dict = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
    dataframes = [df for df in sheet_dict.values()]
    combined_df = pd.concat(dataframes, ignore_index=True) if dataframes else pd.DataFrame()
    if 'Sl No.' in combined_df.columns:
        combined_df = combined_df.drop(columns=['Sl No.'])
    return combined_df

def universal_fuzzy_filter(data, search_value, threshold=80):
    result = pd.DataFrame()
    
    # Ensure search_value is a string
    search_value = str(search_value) if search_value else ""
    
    for column in data.select_dtypes(include=['object', 'string']).columns:
        column_data = data[column].dropna().astype(str)  # Convert all values to strings and drop NaNs
        matched_values = process.extract(
            search_value, column_data.unique(), limit=len(column_data), scorer=fuzz.partial_ratio
        )
        
        # Filter matched values based on threshold
        matched_values = [match[0] for match in matched_values if match[1] >= threshold]
        
        # If there are no matches above the threshold, skip this column
        if not matched_values:
            continue
        
        # Filter rows based on the matched values
        filtered = data[data[column].isin(matched_values)]
        result = pd.concat([result, filtered])
    
    return result.drop_duplicates()


def highlight_search(s, query):
    query = str(query).lower()  # Convert query to lowercase string
    return ['background-color: yellow' if query in str(val).lower() else '' for val in s]


st.set_page_config(page_title="LinkView360", layout="wide")

st.sidebar.title("Login")
username = st.sidebar.text_input("Username")
password = st.sidebar.text_input("Password", type="password")

correct_username = "admin"
correct_password = "password123"

# Check login
if username == correct_username and password == correct_password:
    st.sidebar.success("Login successful!")
    
    # File upload section on separate page
    st.sidebar.title("File Upload")
    uploaded_file = st.sidebar.file_uploader("Upload an Excel file", type=['xlsx'])
    if uploaded_file is not None:
        file_path = os.path.join(EXCEL_FOLDER, uploaded_file.name)
        with open(file_path, 'wb') as f:
            f.write(uploaded_file.getbuffer())
        st.sidebar.success(f"File '{uploaded_file.name}' uploaded successfully!")

st.title("LinkView360: Excel Search")
excel_files = load_excel_files()

file_choice = st.selectbox("Select an Excel file", excel_files)

filtered_data = pd.DataFrame()

if file_choice:
    file_path = os.path.join(EXCEL_FOLDER, file_choice)
    data = load_all_sheets(file_path)
    
    if not data.empty:
        st.write("Data Preview (without 'Sl No.')")
        st.dataframe(data.head(100))

        if st.checkbox("Drop rows with missing values"):
            data = data.dropna()

        search_value = st.text_input("Enter a value to search (fuzzy matching across text columns only)")
        
        if search_value:
            filtered_data = universal_fuzzy_filter(data, search_value)
            
            if not filtered_data.empty:
                st.write("Filtered Results:")
                st.dataframe(filtered_data.style.apply(lambda s: highlight_search(s, search_value), axis=1))

                if 'Box Link' in filtered_data.columns:
                    st.write("Box Links:")
                    for index, row in filtered_data.iterrows():
                        st.write(f"[{row['Box Link']}]({row['Box Link']})")
            else:
                st.warning(f"No results found for '{search_value}'.")
        else:
            st.warning("Please enter a value to search.")

        if not filtered_data.empty:
            @st.cache_data
            def convert_df(df):
                return df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="Download filtered data as CSV",
                data=convert_df(filtered_data),
                file_name='filtered_data.csv',
                mime='text/csv',
            )
        if st.button("Reset Search"):
            st.experimental_rerun()
else:
    st.warning("Please login and upload an Excel file to proceed.")
