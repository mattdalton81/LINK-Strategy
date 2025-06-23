
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from wordcloud import WordCloud
from collections import Counter
import re

# Load Excel file
@st.cache_data
def load_data(sheet_name):
    df = pd.read_excel("Strategy Day responses.xlsx", sheet_name=sheet_name, engine="openpyxl")
    return df

# Preprocess text: combine all text into one string and clean it
def preprocess_text(df):
    text = " ".join(df.astype(str).fillna("").values.flatten())
    text = re.sub(r"[^a-zA-Z0-9\s]", "", text)
    return text.lower()

# Generate word cloud
def generate_wordcloud(text):
    wordcloud = WordCloud(width=800, height=400, background_color="white").generate(text)
    return wordcloud

# Generate word frequency bar chart
def plot_word_frequencies(text, top_n=20):
    words = text.split()
    word_counts = Counter(words)
    common_words = word_counts.most_common(top_n)
    words, counts = zip(*common_words)
    fig, ax = plt.subplots()
    ax.barh(words[::-1], counts[::-1])
    ax.set_xlabel("Frequency")
    ax.set_title("Top Word Frequencies")
    return fig

# Main app
st.title("Strategy Day Feedback Dashboard")

# Load sheet names
xls = pd.ExcelFile("Strategy Day responses.xlsx", engine="openpyxl")
sheet_names = xls.sheet_names

# Sidebar for sheet selection
selected_sheet = st.sidebar.selectbox("Select a sheet", sheet_names)

# Load and display data
df = load_data(selected_sheet)
st.subheader(f"Raw Data - {selected_sheet}")
st.dataframe(df)

# Preprocess and visualize
text = preprocess_text(df)

st.subheader("Word Cloud")
wordcloud = generate_wordcloud(text)
st.image(wordcloud.to_array())

st.subheader("Top Word Frequencies")
fig = plot_word_frequencies(text)
st.pyplot(fig)
