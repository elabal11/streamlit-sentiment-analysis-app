import streamlit as st
import pandas as pd
from textblob import TextBlob
import io
import xlsxwriter
import altair as alt

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

#### This app:

- Provides a template Excel file to be filled with a column of text comments.
- Accepts an uploaded Excel file with the filled comments.
- Performs sentiment analysis on each comment.
- Allows you to download an Excel file with the results.

You can download the template file, fill it with your text comments, and then upload it back to the app.

#### The sentiment analysis calculates:

1. **Polarity**: This is a measure that lies in the range of [-1,1]. A value closer to -1 means that the sentiment is more negative, a value closer to +1 means the sentiment is more positive, and a value around 0 indicates a neutral sentiment.
2. **Subjectivity**: This is a measure that lies in the range of [0,1]. A value closer to 0 means that the text is more objective, while a value closer to 1 means the text is more subjective.
''')

st.subheader('Input')

# Download button for the template file
with open("template_sentiment_analysis.xlsx", "rb") as f:
    bytes_data = f.read()
st.download_button(
    label="Download Template Excel File",
    data=bytes_data,
    file_name="template_sentiment_analysis.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

uploaded_file = st.file_uploader('Please upload your filled excel file below', type=['xlsx'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    with st.spinner("Processing..."):
        df_processed = perform_sentiment_analysis(df)
        st.success("Processing complete!")

        # Prepare processed data for download
        st.subheader('Output')
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

        st.subheader('Visualisation')
        
        # Altair scatter plot
        st.altair_chart(alt.Chart(df_processed).mark_circle(size=60).encode(x='polarity', y='subjectivity', 
                        tooltip=['text', 'polarity', 'subjectivity']).interactive(), use_container_width=True)
        
        # Altair histograms
        st.altair_chart(alt.Chart(df_processed).mark_bar().encode(
            alt.X("polarity", bin=True),
            y='count()',).properties(width=300, height=200), use_container_width=True)

        st.altair_chart(alt.Chart(df_processed).mark_bar().encode(
            alt.X("subjectivity", bin=True),
            y='count()',).properties(width=300, height=200), use_container_width=True)

        # Display selected number of rows with a slider
        row_number = st.slider('Select number of rows to preview', min(10,len(df_processed)), len(df_processed))
        st.dataframe(df_processed.head(row_number))
