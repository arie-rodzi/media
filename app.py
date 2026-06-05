
import streamlit as st
import pandas as pd
import feedparser
from urllib.parse import quote
from datetime import datetime, timedelta, date
import re
from io import BytesIO

# =========================================================
# SPiKOM Premium - Sistem Pemantauan Isu Komuniti Malaysia
# Fokus: Tarik berita automatik, cari isu dominan, negeri/daerah,
# contoh berita, export Excel. Tanpa sidebar.
# =========================================================

st.set_page_config(
    page_title="SPiKOM Premium",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ---------------- CSS PREMIUM ----------------
st.markdown("""
<style>
[data-testid="stSidebar"] {display: none;}
[data-testid="collapsedControl"] {display: none;}
.block-container {
    padding-top: 1.2rem;
    padding-left: 2.5rem;
    padding-right: 2.5rem;
    max-width: 1500px;
}
.stApp {
    background:
    radial-gradient(circle at top left, rgba(255, 214, 102, 0.22), transparent 28%),
    radial-gradient(circle at top right, rgba(0, 44, 95, 0.25), transparent 30%),
    linear-gradient(135deg, #f8fafc 0%, #eef3f8 45%, #fdf7e3 100%);
}
.hero {
    padding: 30px 34px;
    border-radius: 28px;
    background: linear-gradient(135deg, #001f3f 0%, #003b73 55%, #c99a2e 100%);
    color: white;
    box-shadow: 0 22px 55px rgba(0, 31, 63, 0.25);
    margin-bottom: 22px;
}
.hero h1 {
    font-size: 2.3rem;
    font-weight: 850;
    margin-bottom: 6px;
}
.hero p {
    font-size: 1.02rem;
    opacity: 0.96;
    margin: 0;
}
.card {
    border-radius: 22px;
    padding: 22px;
    background: rgba(255,255,255,0.86);
    border: 1px solid rgba(0,31,63,0.08);
    box-shadow: 0 14px 35px rgba(15, 23, 42, 0.08);
    margin-bottom: 18px;
}
.kpi {
    border-radius: 22px;
    padding: 20px;
    background: linear-gradient(180deg, #ffffff 0%, #f8fbff 100%);
    border: 1px solid rgba(0, 31, 63, 0.08);
    box-shadow: 0 12px 30px rgba(15, 23, 42, 0.08);
    min-height: 120px;
}
.kpi .label {
    color: #64748b;
    font-size: 0.92rem;
    font-weight: 650;
}
.kpi .value {
    color: #001f3f;
    font-size: 2.15rem;
    font-weight: 900;
    margin-top: 8px;
}
.badge {
    display:inline-block;
    padding: 7px 12px;
    border-radius: 999px;
    background: #fff4cc;
    color: #745000;
    font-weight: 750;
    font-size: 0.83rem;
    margin: 4px 4px 4px 0;
}
.section-title {
    color: #001f3f;
    font-weight: 850;
    margin-top: 12px;
}
.newsbox {
    padding: 15px 18px;
    border-left: 5px solid #c99a2e;
    border-radius: 14px;
    background: #ffffff;
    margin-bottom: 10px;
    box-shadow: 0 8px 20px rgba(15,23,42,0.06);
}
.newsbox a {
    color: #003b73;
    font-weight: 750;
    text-decoration: none;
}
div.stButton > button {
    width: 100%;
    border-radius: 18px;
    padding: 0.8rem 1rem;
    font-weight: 850;
    background: linear-gradient(135deg, #001f3f, #00509d);
    color: white;
    border: none;
    box-shadow: 0 12px 25px rgba(0,31,63,0.22);
}
div.stDownloadButton > button {
    border-radius: 18px;
    padding: 0.7rem 1rem;
    font-weight: 800;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
    <h1>📡 SPiKOM Premium</h1>
    <p>Sistem Pemantauan Isu Komuniti Malaysia — kutipan berita automatik, pengesanan isu dominan, negeri/daerah, contoh berita dan eksport data.</p>
</div>
""", unsafe_allow_html=True)

# ---------------- KEYWORD BANK ----------------
KEYWORD_BANK = {
    "Kaum & Perpaduan": [
        "perpaduan Malaysia",
        "isu kaum Malaysia",
        "perkauman Malaysia",
        "diskriminasi kaum Malaysia",
        "hubungan kaum Malaysia",
        "keharmonian kaum Malaysia",
        "sensitiviti kaum Malaysia",
        "komuniti majmuk Malaysia"
    ],
    "Agama & Sensitiviti": [
        "isu agama Malaysia",
        "sensitiviti agama Malaysia",
        "konflik agama Malaysia",
        "keharmonian agama Malaysia",
        "rumah ibadat Malaysia",
        "penghinaan agama Malaysia",
        "toleransi agama Malaysia"
    ],
    "Politik & Tadbir Urus": [
        "politik Malaysia",
        "ketegangan politik Malaysia",
        "isu kerajaan Malaysia",
        "parlimen Malaysia",
        "pilihan raya Malaysia",
        "parti politik Malaysia",
        "integriti kerajaan Malaysia",
        "kepercayaan institusi Malaysia"
    ],
    "Ekonomi & Kos Sara Hidup": [
        "kos sara hidup Malaysia",
        "harga barang Malaysia",
        "harga makanan Malaysia",
        "inflasi Malaysia",
        "pengangguran Malaysia",
        "kemiskinan bandar Malaysia",
        "bantuan rakyat Malaysia",
        "gaji Malaysia"
    ],
    "Sosial & Keselamatan Komuniti": [
        "jenayah Malaysia",
        "buli Malaysia",
        "masalah sosial Malaysia",
        "keselamatan komuniti Malaysia",
        "vandalisme Malaysia",
        "dadah Malaysia",
        "remaja Malaysia",
        "rumah tangga Malaysia"
    ],
    "Digital & Maklumat Palsu": [
        "berita palsu Malaysia",
        "fitnah media sosial Malaysia",
        "scam Malaysia",
        "penipuan online Malaysia",
        "ujaran kebencian Malaysia",
        "provokasi media sosial Malaysia",
        "viral Malaysia"
    ],
    "Institusi & Perkhidmatan Awam": [
        "isu polis Malaysia",
        "isu mahkamah Malaysia",
        "isu hospital Malaysia",
        "isu sekolah Malaysia",
        "isu pihak berkuasa tempatan Malaysia",
        "aduan rakyat Malaysia",
        "perkhidmatan awam Malaysia"
    ]
}

ISSUE_THEMES = {
    "Kos sara hidup": ["kos sara hidup", "harga barang", "harga makanan", "inflasi", "subsidi", "gaji", "bantuan rakyat"],
    "Politik": ["politik", "parti", "parlimen", "menteri", "kerajaan", "pilihan raya", "dun", "undi"],
    "Jenayah & keselamatan": ["jenayah", "bunuh", "rompak", "curi", "buli", "keselamatan", "polis", "mahkamah"],
    "Agama": ["agama", "rumah ibadat", "masjid", "gereja", "tokong", "kuil", "sensitiviti agama"],
    "Kaum & perpaduan": ["kaum", "perkauman", "perpaduan", "diskriminasi", "komuniti majmuk"],
    "Digital & scam": ["scam", "online", "media sosial", "viral", "fitnah", "berita palsu", "ujaran kebencian"],
    "Institusi awam": ["hospital", "sekolah", "pbt", "agensi", "perkhidmatan awam", "kerajaan tempatan"],
    "Sosial komuniti": ["remaja", "keluarga", "komuniti", "masalah sosial", "vandalisme", "penduduk"]
}

STATES = [
    "Johor", "Kedah", "Kelantan", "Melaka", "Negeri Sembilan", "Pahang", "Pulau Pinang",
    "Perak", "Perlis", "Sabah", "Sarawak", "Selangor", "Terengganu",
    "Kuala Lumpur", "Putrajaya", "Labuan"
]

DISTRICTS = [
    "Seremban", "Port Dickson", "Jempol", "Tampin", "Kuala Pilah", "Rembau", "Jelebu",
    "Petaling", "Shah Alam", "Klang", "Gombak", "Hulu Langat", "Sepang", "Kajang",
    "Johor Bahru", "Batu Pahat", "Muar", "Kluang", "Kota Bharu", "Kuantan", "Kota Kinabalu",
    "Kuching", "Sibu", "Miri", "Ipoh", "Taiping", "Alor Setar", "Sungai Petani",
    "George Town", "Butterworth", "Kuala Terengganu", "Kangar", "Melaka Tengah"
]

# ---------------- FUNCTIONS ----------------
def parse_google_date(x):
    try:
        return pd.to_datetime(x, utc=True).tz_convert(None)
    except Exception:
        return pd.NaT

def find_location(text, location_list):
    if not isinstance(text, str):
        return ""
    found = []
    low = text.lower()
    for loc in location_list:
        if loc.lower() in low:
            found.append(loc)
    return ", ".join(sorted(set(found)))

def classify_issue(text):
    low = text.lower() if isinstance(text, str) else ""
    scores = {}
    for theme, words in ISSUE_THEMES.items():
        scores[theme] = sum(1 for w in words if w.lower() in low)
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else "Lain-lain"

def get_google_news(keyword, start_date, end_date, max_items=100):
    # Google News accepts after/before in query. before is exclusive, add one day.
    after = start_date.strftime("%Y-%m-%d")
    before = (end_date + timedelta(days=1)).strftime("%Y-%m-%d")
    full_query = f'{keyword} after:{after} before:{before}'
    q = quote(full_query)
    url = f"https://news.google.com/rss/search?q={q}&hl=ms-MY&gl=MY&ceid=MY:ms"

    feed = feedparser.parse(url)
    rows = []
    for e in feed.entries[:max_items]:
        title = e.get("title", "")
        published = e.get("published", "")
        source = ""
        try:
            source = e.get("source", {}).get("title", "")
        except Exception:
            source = ""
        rows.append({
            "Tarikh Kutipan": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Keyword": keyword,
            "Tajuk Berita": title,
            "Sumber": source if source else "Google News",
            "Tarikh Berita Raw": published,
            "Pautan": e.get("link", ""),
            "RSS Query": full_query
        })
    return rows

def to_excel_bytes(df, summary_issue, summary_state, summary_district):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Data Berita")
        summary_issue.to_excel(writer, index=False, sheet_name="Top Isu")
        summary_state.to_excel(writer, index=False, sheet_name="Top Negeri")
        summary_district.to_excel(writer, index=False, sheet_name="Top Daerah")
    return output.getvalue()

# ---------------- CONTROL PANEL NO SIDEBAR ----------------
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown("### ⚙️ Tetapan Carian Data")

c1, c2, c3, c4 = st.columns([1,1,1,1])
with c1:
    default_start = date.today() - timedelta(days=14)
    start_date = st.date_input("Tarikh Mula", value=default_start)
with c2:
    end_date = st.date_input("Tarikh Tamat", value=date.today())
with c3:
    max_per_keyword = st.slider("Maksimum berita setiap keyword", 20, 100, 80, step=10)
with c4:
    show_examples = st.slider("Contoh berita setiap isu", 3, 10, 5)

selected_dimensions = st.multiselect(
    "Pilih dimensi keyword",
    list(KEYWORD_BANK.keys()),
    default=list(KEYWORD_BANK.keys())
)

all_default_keywords = []
for dim in selected_dimensions:
    all_default_keywords.extend(KEYWORD_BANK[dim])

keyword_text = st.text_area(
    "Keyword carian — boleh edit/tambah sendiri. Satu keyword satu baris.",
    value="\n".join(all_default_keywords),
    height=190
)

st.markdown("</div>", unsafe_allow_html=True)

# ---------------- RUN BUTTON ----------------
if st.button("🔍 Cari Berita & Jana Dashboard"):
    if start_date > end_date:
        st.error("Tarikh mula tidak boleh melebihi tarikh tamat.")
        st.stop()

    keywords = [k.strip() for k in keyword_text.splitlines() if k.strip()]
    if not keywords:
        st.error("Sila masukkan sekurang-kurangnya satu keyword.")
        st.stop()

    progress = st.progress(0)
    all_rows = []
    for i, kw in enumerate(keywords):
        all_rows.extend(get_google_news(kw, start_date, end_date, max_per_keyword))
        progress.progress((i + 1) / len(keywords))

    df = pd.DataFrame(all_rows)

    if df.empty:
        st.warning("Tiada berita ditemui. Cuba longgarkan keyword atau panjangkan tempoh tarikh.")
        st.stop()

    # Cleaning
    df["Tarikh Berita"] = df["Tarikh Berita Raw"].apply(parse_google_date)
    df = df.drop_duplicates(subset=["Tajuk Berita"]).reset_index(drop=True)

    # Date filter after parsing too
    sd = pd.to_datetime(start_date)
    ed = pd.to_datetime(end_date) + pd.Timedelta(days=1)
    df = df[(df["Tarikh Berita"].isna()) | ((df["Tarikh Berita"] >= sd) & (df["Tarikh Berita"] < ed))]

    df["Isu Dikesan"] = df["Tajuk Berita"].apply(classify_issue)
    df["Negeri Dikesan"] = df["Tajuk Berita"].apply(lambda x: find_location(x, STATES))
    df["Daerah Dikesan"] = df["Tajuk Berita"].apply(lambda x: find_location(x, DISTRICTS))

    df["Tarikh Berita"] = df["Tarikh Berita"].dt.strftime("%Y-%m-%d %H:%M:%S")
    df = df[[
        "Tarikh Kutipan", "Tarikh Berita", "Isu Dikesan", "Negeri Dikesan", "Daerah Dikesan",
        "Keyword", "Sumber", "Tajuk Berita", "Pautan", "RSS Query"
    ]]

    # Summary
    summary_issue = df["Isu Dikesan"].value_counts().reset_index()
    summary_issue.columns = ["Isu", "Bilangan Berita"]

    summary_state = df[df["Negeri Dikesan"] != ""]["Negeri Dikesan"].value_counts().reset_index()
    summary_state.columns = ["Negeri", "Bilangan Berita"]

    summary_district = df[df["Daerah Dikesan"] != ""]["Daerah Dikesan"].value_counts().reset_index()
    summary_district.columns = ["Daerah", "Bilangan Berita"]

    top_issue = summary_issue.iloc[0]["Isu"] if len(summary_issue) else "Tiada"
    top_state = summary_state.iloc[0]["Negeri"] if len(summary_state) else "Belum jelas"
    top_district = summary_district.iloc[0]["Daerah"] if len(summary_district) else "Belum jelas"

    # KPI
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(f'<div class="kpi"><div class="label">Jumlah Berita Unik</div><div class="value">{len(df)}</div></div>', unsafe_allow_html=True)
    with k2:
        st.markdown(f'<div class="kpi"><div class="label">Isu Paling Tinggi</div><div class="value" style="font-size:1.35rem;">{top_issue}</div></div>', unsafe_allow_html=True)
    with k3:
        st.markdown(f'<div class="kpi"><div class="label">Negeri Paling Disebut</div><div class="value" style="font-size:1.35rem;">{top_state}</div></div>', unsafe_allow_html=True)
    with k4:
        st.markdown(f'<div class="kpi"><div class="label">Daerah Paling Disebut</div><div class="value" style="font-size:1.35rem;">{top_district}</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 📊 Top Isu Dikesan")
    st.bar_chart(summary_issue.set_index("Isu"))
    st.dataframe(summary_issue, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    cA, cB = st.columns(2)
    with cA:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("### 🗺️ Negeri Paling Banyak Disebut")
        if not summary_state.empty:
            st.bar_chart(summary_state.set_index("Negeri"))
            st.dataframe(summary_state, use_container_width=True)
        else:
            st.info("Negeri tidak jelas dalam tajuk berita. Perlu baca kandungan penuh untuk ketepatan lebih tinggi.")
        st.markdown("</div>", unsafe_allow_html=True)

    with cB:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("### 📍 Daerah Paling Banyak Disebut")
        if not summary_district.empty:
            st.bar_chart(summary_district.set_index("Daerah"))
            st.dataframe(summary_district, use_container_width=True)
        else:
            st.info("Daerah tidak jelas dalam tajuk berita. Boleh tambah senarai daerah dalam kod.")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 📰 Data Berita Lengkap")
    st.dataframe(df, use_container_width=True, height=450)
    st.markdown("</div>", unsafe_allow_html=True)

    # Example news by issue
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 🧾 Contoh Berita Mengikut Isu")
    for issue in summary_issue["Isu"].head(10):
        sub = df[df["Isu Dikesan"] == issue].head(show_examples)
        st.markdown(f"#### {issue} ({len(df[df['Isu Dikesan'] == issue])} berita)")
        for _, r in sub.iterrows():
            title = str(r["Tajuk Berita"])
            link = str(r["Pautan"])
            source = str(r["Sumber"])
            tarikh = str(r["Tarikh Berita"])
            negeri = str(r["Negeri Dikesan"]) if str(r["Negeri Dikesan"]) else "Tidak jelas"
            daerah = str(r["Daerah Dikesan"]) if str(r["Daerah Dikesan"]) else "Tidak jelas"
            st.markdown(
                f"""
                <div class="newsbox">
                    <a href="{link}" target="_blank">{title}</a><br>
                    <small><b>Sumber:</b> {source} &nbsp; | &nbsp; <b>Tarikh:</b> {tarikh} &nbsp; | &nbsp; <b>Negeri:</b> {negeri} &nbsp; | &nbsp; <b>Daerah:</b> {daerah}</small>
                </div>
                """,
                unsafe_allow_html=True
            )
    st.markdown("</div>", unsafe_allow_html=True)

    # Executive summary template
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 📝 Rumusan Automatik Asas")
    st.write(
        f"Dalam tempoh **{start_date} hingga {end_date}**, sistem mengesan **{len(df)} berita unik** "
        f"berdasarkan {len(keywords)} keyword. Isu paling dominan ialah **{top_issue}**. "
        f"Negeri yang paling banyak disebut ialah **{top_state}**, manakala daerah paling banyak disebut ialah **{top_district}**. "
        "Nota: Rumusan ini berasaskan tajuk berita. Untuk analisis lebih tepat, versi seterusnya boleh membaca kandungan artikel penuh dan menggunakan AI clustering."
    )
    st.markdown("</div>", unsafe_allow_html=True)

    # Downloads
    excel_bytes = to_excel_bytes(df, summary_issue, summary_state, summary_district)
    st.download_button(
        "⬇️ Download Excel Lengkap",
        data=excel_bytes,
        file_name=f"SPiKOM_Data_{start_date}_hingga_{end_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### Cara guna")
    st.write("1. Pilih tarikh mula dan tarikh tamat.")
    st.write("2. Pilih dimensi keyword atau tambah keyword sendiri.")
    st.write("3. Tekan **Cari Berita & Jana Dashboard**.")
    st.write("4. Sistem akan keluarkan jumlah berita, isu paling tinggi, negeri/daerah disebut, contoh berita dan fail Excel.")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### Nota tentang media sosial")
    st.write(
        "Versi ini fokus kepada Google News RSS kerana paling mudah dan stabil untuk prototaip. "
        "Facebook, Instagram dan Threads memerlukan Meta Graph API/Threads API serta permission rasmi. "
        "X/Twitter pula menggunakan API rasmi berbayar/pay-per-use. Untuk proposal, nyatakan media sosial sebagai fasa integrasi kedua."
    )
    st.markdown("</div>", unsafe_allow_html=True)
