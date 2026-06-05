
import streamlit as st
import pandas as pd
import feedparser
from urllib.parse import quote
from datetime import date

st.set_page_config(page_title="SPiKOM", layout="wide")

st.title("SPiKOM - Sistem Pemantauan Isu Komuniti Malaysia")

start_date = st.date_input("Tarikh Mula")
end_date = st.date_input("Tarikh Tamat")

keywords = st.text_area(
    "Keyword (satu baris satu keyword)",
    "perpaduan Malaysia\nisu agama Malaysia\npolitik Malaysia\nkos sara hidup Malaysia"
)

def get_news(keyword):
    q = quote(keyword)
    url = f"https://news.google.com/rss/search?q={q}&hl=ms&gl=MY&ceid=MY:ms"
    feed = feedparser.parse(url)
    rows = []
    for e in feed.entries:
        rows.append({
            "Keyword": keyword,
            "Tajuk": e.get("title",""),
            "Tarikh": e.get("published",""),
            "Pautan": e.get("link","")
        })
    return rows

if st.button("Cari Berita"):
    all_rows = []
    for kw in keywords.splitlines():
        kw = kw.strip()
        if kw:
            all_rows.extend(get_news(kw))

    df = pd.DataFrame(all_rows).drop_duplicates(subset=["Tajuk"])

    st.metric("Jumlah Berita", len(df))
    st.dataframe(df, use_container_width=True)

    csv = df.to_csv(index=False).encode()
    st.download_button(
        "Download CSV",
        csv,
        file_name="spikom_news.csv",
        mime="text/csv"
    )

    st.subheader("Contoh Ringkasan Isu")
    st.info("Versi ini prototaip. Boleh ditambah AI, PDF, Word Cloud dan pengelasan isu.")
