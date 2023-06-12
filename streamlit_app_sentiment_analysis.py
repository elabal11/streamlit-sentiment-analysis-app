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
## Description

This app:

- Provides a template Excel file to be filled with a column of text comments.
- Accepts an uploaded Excel file with the filled comments.
- Performs sentiment analysis on each comment.
- Allows you to download an Excel file with the results.

You can download the template file, fill it with your text comments, and then upload it back to the app.

### Sentiment Analysis

The sentiment analysis calculates:

1. **Polarity**: This is a measure that lies in the range of [-1,1]. A value closer to -1 means that the sentiment is more negative, a value closer to +1 means the sentiment is more positive, and a value around 0 indicates a neutral sentiment.
2. **Subjectivity**: This is a measure that lies in the range of [0,1]. A value closer to 0 means that the text is more objective, while a value closer to 1 means the text is more subjective.
''')

# Download button for the template file
with open("template_sentiment_analysis.xlsx", "rb") as f:
    bytes_data = f.read()
st.download_button(
    label="Download Template File",
    data=bytes_data,
    file_name="template_sentiment_analysis.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

uploaded_file = st.file_uploader('Please upload your filled excel file', type=['xlsx'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df_processed = perform_sentiment_analysis(df)

    # Prepare processed data for download
    towrite = io.BytesIO()
    df_processed.to_excel(towrite, engine='xlsxwriter', index=False, sheet_name='Sheet1')
    towrite.seek(0)  # reset pointer
    downloaded_file = towrite.read()  # read data to a variable

    # Use st.download_button to create a download button for processed data
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
