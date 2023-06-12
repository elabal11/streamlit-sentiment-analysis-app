import streamlit as st
import pandas as pd
from textblob import TextBlob
import io
import xlsxwriter
import matplotlib.pyplot as plt

# Function to perform sentiment analysis
def perform_sentiment_analysis(df):
    df['polarity'] = df['text'].apply(lambda x: TextBlob(x).sentiment.polarity)
    df['subjectivity'] = df['text'].apply(lambda x: TextBlob(x).sentiment.subjectivity)
    return df

# App title
st.title('Sentiment Analysis App')

# App description
st.markdown('''
This app accepts an Excel file with a column of text comments, performs sentiment analysis on each comment, 
and allows you to download an Excel file with the results. The sentiment analysis calculates the polarity and subjectivity of each comment. 
Polarity is a measure of the positivity or negativity of the comment. Subjectivity is a measure of the objectivity or subjectivity of the comment.
''')

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

    # Plot histograms for polarity and subjectivity
    fig, ax = plt.subplots(2, 1, figsize=(12, 12))

    ax[0].hist(df_processed['polarity'], bins=20, color='blue', alpha=0.7)
    ax[0].set_title('Polarity Histogram')

    ax[1].hist(df_processed['subjectivity'], bins=20, color='green', alpha=0.7)
    ax[1].set_title('Subjectivity Histogram')

    plt.tight_layout()
    st.pyplot(fig)

    # Display selected number of rows with a slider
    row_number = st.slider('Select number of rows to view', 1, len(df_processed))
    st.dataframe(df_processed.head(row_number))
