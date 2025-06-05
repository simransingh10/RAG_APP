import streamlit as st
import pandas as pd
import requests
import time
import io

# Function to generate description using local LLM via Ollama
def generate_description(column_name, data_type=None, dax_expression=None):
    if pd.notnull(dax_expression) and dax_expression.strip() != "":
        prompt = f"Generate a user-friendly description for the measure named '{column_name}' with the following DAX expression:\n{dax_expression}"
    else:
        prompt = f"Generate a user-friendly description for the column named '{column_name}' with data type '{data_type}'."

    try:
        response = requests.post(
            "http://localhost:11434/api/generate",
            json={
                "model": "llama2",
                "prompt": prompt,
                "stream": False
            }
        )
        if response.status_code == 200:
            result = response.json()
            return result.get("response", "").strip()
        else:
            return f"Error: {response.status_code} - {response.text}"
    except Exception as e:
        return f"Exception: {str(e)}"

# Streamlit UI
st.title("Power BI Metadata Description Generator")
st.write("Upload an Excel file containing your Power BI metadata to generate Copilot-friendly descriptions.")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        required_columns = {"TableName", "ColumnName", "DataType", "DAXExpression"}
        if not required_columns.issubset(df.columns):
            st.error(f"The uploaded file must contain the following columns: {', '.join(required_columns)}")
        else:
            st.success("File uploaded successfully!")
            st.subheader("Uploaded Data")
            st.dataframe(df)

            if st.button("Generate Descriptions"):
                progress_bar = st.progress(0)
                progress_text = st.empty()
                descriptions = []
                total_rows = len(df)
                for index, row in df.iterrows():
                    description = generate_description(
                        column_name=row["ColumnName"],
                        data_type=row.get("DataType", ""),
                        dax_expression=row.get("DAXExpression", "")
                    )
                    descriptions.append(description)
                    percent_complete = int((index + 1) / total_rows * 100)
                    progress_bar.progress(percent_complete)
                    progress_text.text(f"Processing: {percent_complete}% complete")
                    time.sleep(0.1)  # To simulate processing time
                df["Description"] = descriptions
                st.success("Descriptions generated successfully!")
                st.subheader("Data with Descriptions")
                st.dataframe(df)
                # Option to download the result as Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Descriptions')
                processed_data = output.getvalue()
                st.download_button(
                    label="Download Descriptions as Excel",
                    data=processed_data,
                    file_name="metadata_with_descriptions.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"An error occurred while processing the file: {str(e)}")
