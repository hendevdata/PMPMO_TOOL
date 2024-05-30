import streamlit as st
import pandas as pd
import openai

# Set up OpenAI API key
api_key = "your api key"
openai.api_key = api_key

def read_all_sheets_from_excel(file):
    try:
        # Read the Excel file
        xls = pd.ExcelFile(file)
        sheets = {}
        for sheet_name in xls.sheet_names:
            sheets[sheet_name] = xls.parse(sheet_name)
        return sheets
    except UnicodeDecodeError as e:
        st.error(f"Error decoding file: {e}")
        st.write("The file might contain characters that cannot be processed. Try the following:")
        st.write("- Clean the data or use a different file.")
        st.write("- Open the file in a hex editor to inspect the specific byte causing the error.")
        st.stop()

def get_insights(prompt):
    response = openai.Completion.create(
        model="gpt-3.5-turbo-instruct",
        prompt=prompt,
        max_tokens=500
    )
    return response["choices"][0]["text"].strip()


# Streamlit app
st.title("Project Management Portfolio Insights")

st.write("Upload an Excel file to get insights from ChatGPT.")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Read the Excel file
    try:
        dataframes = read_all_sheets_from_excel(uploaded_file)
    except UnicodeDecodeError as e:
        st.error(f"Error decoding file: {e}")
        st.write("The file might contain characters that cannot be processed. Try the following:")
        st.write("- Clean the data or use a different file.")
        st.write("- Open the file in a hex editor to inspect the specific byte causing the error.")
        st.stop()

    # Display the sheet names
    if dataframes:  # Check if dataframes exist after potential error handling
        st.write("Sheets in the uploaded file:")
        sheet_names = list(dataframes.keys())
        st.write(sheet_names)

        # Select a sheet to display
        selected_sheet = st.selectbox("Select a sheet to display", sheet_names)

        if selected_sheet:
            st.write(f"Displaying the first 15 rows of the {selected_sheet} sheet:")
            st.write(dataframes[selected_sheet].head(15))

            # Chat interface
            st.write("Chat with ChatGPT about the data:")
            user_question = st.text_input("Ask a question about the data")

            if user_question:
                prompt = f"You are an expert in project management portfolio. Provide insights based on the following data:\n\n"
                prompt += dataframes[selected_sheet].head(15).to_string()
                prompt += f"\n\nQuestion: {user_question}"
                insights = get_insights(prompt)
                st.write("Insights from ChatGPT:")
                st.write(insights)
