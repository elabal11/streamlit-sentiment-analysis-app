import streamlit as st
import pandas as pd
from textblob import TextBlob
import io

# Function to perform sentiment analysis
def perform_sentiment_analysis(df):
    df['polarity'] = df['text'].apply(lambda x: TextBlob(x).sentiment.polarity)
    df['subjectivity'] = df['text'].apply(lambda x: TextBlob(x).sentiment.subjectivity)
    return df

st.title('Sentiment Analysis App')

uploaded_file = st.file_uploader('Please upload your excel file', type=['xlsx'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df_processed = perform_sentiment_analysis(df)

    # Prepare processed data for download
    towrite = io.BytesIO()
    df_processed.to_excel(towrite, engine='xlsxwriter', index=False, sheet_name='Sheet1')
    towrite.seek(0)  # reset pointer
    downloaded_file = towrite.read()  # read data to a variable

    # Use st.download_button to create a download button
    st.download_button(
        label="Download Processed Data",
        data=downloaded_file,
        file_name="processed_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
