import streamlit as st
import pandas as pd
import plotly.express as px
import google.generativeai as genai
import pydeck as pdk
import requests
from io import BytesIO
import re

st.set_page_config(page_title="Dashboard Penjualan TM & TT", layout="wide")

# =========================
# CSS FIXED HEADER + KPI TERPISAH
# =========================
st.markdown("""
<style>
[data-testid="stAppViewContainer"] .main .block-container {
    padding-top: 100px;
    padding-left: 2rem;
    padding-right: 2rem;
}

/* HEADER FIXED */
.fixed-header {
    position: fixed;
    top: 1.5rem;
    left: calc(21rem + 1rem);
    right: 1rem;
    background: white;
    z-index: 9999;
    padding: 14px 24px;
    border-bottom: 1px solid #ddd;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    border-radius: 0 0 10px 10px;
}

/* KPI FIXED */
.fixed-kpi {
    position: fixed;
    top: 160px;
    left: calc(21rem + 1rem);
    right: 1rem;
    background: white;
    z-index: 9998;
    padding: 16px 24px;
    border-bottom: 1px solid #eee;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    border-radius: 0 0 10px 10px;
    min-height: 70px;
}

.kpi-row {
    display: flex;
    justify-content: space-between;
    gap: 24px;
    align-items: center;
}

.kpi-item {
    flex: 1;
    display: flex;
    flex-direction: column;
}

.kpi-title {
    font-size: 18px;
    font-weight: 600;
    color: #64748b;
    line-height: 1.1;
}

.kpi-value {
    font-size: 20px;
    font-weight: 800;
    color: #111827;
    line-height: 1.2;
    margin-top: 2px;
}

@media (max-width: 1200px) {
    .fixed-header, .fixed-kpi {
        left: 1rem !important;
        right: 1rem !important;
    }
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER FIXED
# =========================
st.markdown("""
<div class="fixed-header">
    <h2 style="margin:0;">⚡ Dashboard Penjualan kluster B & I UID Jawa Timur</h2>
</div>
""", unsafe_allow_html=True)

# =========================
# DEFAULT LINK DATA DI SINI
# =========================
DEFAULT_DATA_URL = "https://ptpln365-my.sharepoint.com/:x:/g/personal/irham_tantowi_ptpln365_onmicrosoft_com/IQB7_H2n8ioZTpRDENhpxex0AdomJBITjOdqBCrn7Uo7RZ4?e=TVbXK7"

def convert_google_drive_url(url):
    match = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
    if match:
        file_id = match.group(1)
        return f"https://drive.google.com/uc?export=download&id={file_id}"

    match = re.search(r"id=([a-zA-Z0-9_-]+)", url)
    if match:
        file_id = match.group(1)
        return f"https://drive.google.com/uc?export=download&id={file_id}"

    return url

def convert_sharepoint_url(url):
    if "download=1" in url:
        return url
    if "?" in url:
        return url + "&download=1"
    return url + "?download=1"

def read_excel_from_url(url):
    if "drive.google.com" in url:
        url = convert_google_drive_url(url)
    elif "sharepoint.com" in url or "onedrive" in url:
        url = convert_sharepoint_url(url)

    response = requests.get(
        url,
        headers={"User-Agent": "Mozilla/5.0"},
        allow_redirects=True
    )

    if response.status_code != 200:
        raise Exception(f"Gagal download file. Status code: {response.status_code}")

    return pd.read_excel(BytesIO(response.content), sheet_name="TM", header=1)

# =========================
# SUMBER DATA
# =========================
st.sidebar.header("📁 Sumber Data")

file_upload = st.sidebar.file_uploader(
    "Upload Excel jika ingin data baru",
    type=["xlsx"]
)

if file_upload is not None:
    df = pd.read_excel(file_upload, sheet_name="TM", header=1)
    st.sidebar.success("Sukses file upload manual.")
else:
    df = read_excel_from_url(DEFAULT_DATA_URL)
    st.sidebar.success("Sukses data default dari link.")

df.columns = df.columns.astype(str).str.strip()

# =========================
# PERIODE DATA
# =========================
bulan_map = {
    "Januari": "Jan",
    "Februari": "Feb",
    "Maret": "Mar",
    "April": "Apr",
    "Mei": "Mei",
    "Juni": "Jun",
    "Juli": "Jul",
    "Agustus": "Agu",
    "September": "Sep",
    "Oktober": "Okt",
    "November": "Nov",
    "Desember": "Des"
}

urutan_bulan = list(bulan_map.values())

st.sidebar.header("🗓️ Periode Data")

tahun_lalu = st.sidebar.number_input("Tahun Lalu", value=2025, step=1)
tahun_ini = st.sidebar.number_input("Tahun Ini", value=2026, step=1)

mode_periode = st.sidebar.selectbox(
    "Mode Periode",
    ["Kumulatif s.d. Bulan", "Bulanan YoY"]
)

pilih_bulan = st.sidebar.selectbox(
    "Pilih Bulan",
    list(bulan_map.keys()),
    index=3
)

kode_bulan = bulan_map[pilih_bulan]
index_bulan = urutan_bulan.index(kode_bulan)

def hitung_kwh(df, tahun, kode_bulan, mode):
    if mode == "Bulanan YoY":
        kolom_bulanan = f"kWh {kode_bulan} {tahun}"
        if kolom_bulanan in df.columns:
            return pd.to_numeric(df[kolom_bulanan], errors="coerce").fillna(0), kolom_bulanan
        return pd.Series(0, index=df.index), f"{kolom_bulanan} belum tersedia"

    kolom_kumulatif = f"kWh sd {kode_bulan} {tahun}"
    if kolom_kumulatif in df.columns:
        return pd.to_numeric(df[kolom_kumulatif], errors="coerce").fillna(0), kolom_kumulatif

    bulan_sampai = urutan_bulan[:index_bulan + 1]
    kolom_bulanan = [
        f"kWh {b} {tahun}"
        for b in bulan_sampai
        if f"kWh {b} {tahun}" in df.columns
    ]

    if len(kolom_bulanan) > 0:
        nilai = df[kolom_bulanan].apply(
            pd.to_numeric, errors="coerce"
        ).fillna(0).sum(axis=1)
        return nilai, " + ".join(kolom_bulanan)

    return pd.Series(0, index=df.index), f"Data {kode_bulan} {tahun} belum tersedia"

nilai_lalu, sumber_lalu = hitung_kwh(df, tahun_lalu, kode_bulan, mode_periode)
nilai_ini, sumber_ini = hitung_kwh(df, tahun_ini, kode_bulan, mode_periode)

st.sidebar.markdown("### 📊 Parameter Data")
st.sidebar.success(f"Mode: {mode_periode}")
st.sidebar.info(f"Sumber {tahun_lalu}: {sumber_lalu}")
st.sidebar.info(f"Sumber {tahun_ini}: {sumber_ini}")

df = df.dropna(subset=["UP3", "TARIF", "KLUSTER USAHA"], how="any")

df["kWh Tahun Lalu"] = nilai_lalu
df["kWh Tahun Ini"] = nilai_ini

df["GWh Tahun Lalu"] = df["kWh Tahun Lalu"] / 1_000_000
df["GWh Tahun Ini"] = df["kWh Tahun Ini"] / 1_000_000
df["Delta GWh"] = df["GWh Tahun Ini"] - df["GWh Tahun Lalu"]

df["Growth %"] = df.apply(
    lambda x: (x["Delta GWh"] / x["GWh Tahun Lalu"] * 100)
    if x["GWh Tahun Lalu"] != 0 else 0,
    axis=1
)

# =========================
# FILTER DATA
# =========================
st.sidebar.header("🔎 Filter Data")

pilih_tarif = st.sidebar.multiselect(
    "Tarif",
    sorted(df["TARIF"].dropna().unique()),
    default=sorted(df["TARIF"].dropna().unique())
)

pilih_up3 = st.sidebar.multiselect(
    "UP3",
    sorted(df["UP3"].dropna().unique()),
    default=sorted(df["UP3"].dropna().unique())
)

pilih_cluster = st.sidebar.multiselect(
    "Cluster KBLI",
    sorted(df["KLUSTER USAHA"].dropna().unique()),
    default=sorted(df["KLUSTER USAHA"].dropna().unique())
)

if "GOLONGAN TARIF" in df.columns:
    pilih_golongan = st.sidebar.multiselect(
        "Golongan Tarif",
        sorted(df["GOLONGAN TARIF"].dropna().unique()),
        default=sorted(df["GOLONGAN TARIF"].dropna().unique())
    )
else:
    pilih_golongan = []

# =========================
# SEARCH IDPEL / PELANGGAN
# =========================
st.sidebar.header("🔍 Cari Pelanggan")

search_idpel = st.sidebar.text_input(
    "Cari IDPEL / Nama Pelanggan",
    placeholder="Masukkan IDPEL atau nama pelanggan"
)

st.sidebar.header("📊 Filter Visual Cluster")

mode_cluster = st.sidebar.selectbox(
    "Mode Tampilan",
    ["Top saja", "Bottom saja", "Top + Bottom"]
)

top_n_cluster = st.sidebar.slider("Jumlah Top Cluster", 3, 30, 10, 1)
bottom_n_cluster = st.sidebar.slider("Jumlah Bottom Cluster", 3, 30, 10, 1)

metrik_cluster = st.sidebar.selectbox(
    "Metrik Ranking",
    ["GWh Tahun Ini", "Delta GWh", "Growth %"]
)

st.sidebar.header("🎚️ Filter Nilai")

kondisi = st.sidebar.selectbox(
    "Kondisi Filter Nilai",
    ["Semua", "Lebih dari", "Kurang dari", "Antara"]
)

nilai_min = 0.0
nilai_max = 0.0

if kondisi == "Lebih dari":
    nilai_min = st.sidebar.number_input("Nilai lebih dari", value=0.0)
elif kondisi == "Kurang dari":
    nilai_max = st.sidebar.number_input("Nilai kurang dari", value=0.0)
elif kondisi == "Antara":
    nilai_min = st.sidebar.number_input("Nilai minimum", value=0.0)
    nilai_max = st.sidebar.number_input("Nilai maksimum", value=1000.0)

st.sidebar.header("🗺️ Map Sebaran")
tampilkan_map = st.sidebar.checkbox("Tampilkan Map Pelanggan", value=False)

st.sidebar.header("🧠 AI Insight")
api_key = st.secrets["GOOGLE_API_KEY"]

# =========================
# APPLY FILTER
# =========================
df_filter = df[
    (df["TARIF"].isin(pilih_tarif)) &
    (df["UP3"].isin(pilih_up3)) &
    (df["KLUSTER USAHA"].isin(pilih_cluster))
]

if "GOLONGAN TARIF" in df.columns:
    df_filter = df_filter[df_filter["GOLONGAN TARIF"].isin(pilih_golongan)]

if kondisi == "Lebih dari":
    df_filter = df_filter[df_filter[metrik_cluster] > nilai_min]
elif kondisi == "Kurang dari":
    df_filter = df_filter[df_filter[metrik_cluster] < nilai_max]
elif kondisi == "Antara":
    df_filter = df_filter[
        (df_filter[metrik_cluster] >= nilai_min) &
        (df_filter[metrik_cluster] <= nilai_max)
    ]

# =========================
# KPI FIXED
# =========================
total_lalu = df_filter["GWh Tahun Lalu"].sum()
total_ini = df_filter["GWh Tahun Ini"].sum()
delta = df_filter["Delta GWh"].sum()
growth = (delta / total_lalu * 100) if total_lalu != 0 else 0

st.markdown(f"""
<div class="fixed-kpi">
    <div class="kpi-row">
        <div class="kpi-item">
            <span class="kpi-title">{tahun_lalu}</span>
            <span class="kpi-value">{total_lalu:,.2f} GWh</span>
        </div>
        <div class="kpi-item">
            <span class="kpi-title">{tahun_ini}</span>
            <span class="kpi-value">{total_ini:,.2f} GWh</span>
        </div>
        <div class="kpi-item">
            <span class="kpi-title">Delta</span>
            <span class="kpi-value">{delta:,.2f} GWh</span>
        </div>
        <div class="kpi-item">
            <span class="kpi-title">Growth</span>
            <span class="kpi-value">{growth:.2f}%</span>
        </div>
    </div>
    <div style="font-size:12px;color:#64748b;margin-top:6px;">
        Periode: {mode_periode} {pilih_bulan} | Perbandingan {tahun_lalu} vs {tahun_ini}
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("<div style='height: 250px;'></div>", unsafe_allow_html=True)
# =========================
# HASIL SEARCH IDPEL / PELANGGAN
# =========================
if search_idpel:
    st.subheader("🔍 Hasil Pencarian Pelanggan")

    df_search = df_filter.copy()

    if "IDPEL" not in df_search.columns:
        st.warning("Kolom IDPEL tidak ditemukan di data.")
    else:
        df_search["IDPEL"] = df_search["IDPEL"].astype(str)

        if "NAMA PELANGGAN" in df_search.columns:
            df_search["NAMA PELANGGAN"] = df_search["NAMA PELANGGAN"].astype(str)

            hasil = df_search[
                df_search["IDPEL"].str.contains(search_idpel, case=False, na=False) |
                df_search["NAMA PELANGGAN"].str.contains(search_idpel, case=False, na=False)
            ]
        else:
            hasil = df_search[
                df_search["IDPEL"].str.contains(search_idpel, case=False, na=False)
            ]

        if hasil.empty:
            st.info("Data pelanggan tidak ditemukan.")
        else:
            kolom_tampil = [
                "UP3",
                "NAMA PELANGGAN",
                "IDPEL",
                "TARIF",
                "DAYA",
                "KLUSTER USAHA",
                "GWh Tahun Lalu",
                "GWh Tahun Ini",
                "Delta GWh",
                "Growth %"
            ]

            kolom_tampil = [kol for kol in kolom_tampil if kol in hasil.columns]

            st.dataframe(
                hasil[kolom_tampil].sort_values("GWh Tahun Ini", ascending=False),
                use_container_width=True
            )

            if len(hasil) == 1:
                row = hasil.iloc[0]

                st.markdown(f"""
                ### 📌 Detail Pelanggan

                **Nama Pelanggan:** {row.get("NAMA PELANGGAN", "-")}  
                **IDPEL:** {row.get("IDPEL", "-")}  
                **UP3:** {row.get("UP3", "-")}  
                **Tarif:** {row.get("TARIF", "-")}  
                **Daya:** {row.get("DAYA", "-")}  
                **Cluster:** {row.get("KLUSTER USAHA", "-")}  

                **Penjualan Tahun Lalu:** {row.get("GWh Tahun Lalu", 0):,.4f} GWh  
                **Penjualan Tahun Ini:** {row.get("GWh Tahun Ini", 0):,.4f} GWh  
                **Delta:** {row.get("Delta GWh", 0):,.4f} GWh  
                **Growth:** {row.get("Growth %", 0):,.2f}%
                """)

# =========================
# MAP SEBARAN PELANGGAN
# =========================
if tampilkan_map:
    st.subheader("🗺️ Map Sebaran Pelanggan / Cluster")

    if "LATITUDE" not in df_filter.columns or "LONGITUDE" not in df_filter.columns:
        st.warning("Kolom LATITUDE dan LONGITUDE belum tersedia di Excel.")
    else:
        map_df = df_filter.copy()

        map_df["LATITUDE"] = pd.to_numeric(map_df["LATITUDE"], errors="coerce")
        map_df["LONGITUDE"] = pd.to_numeric(map_df["LONGITUDE"], errors="coerce")

        map_df = map_df.dropna(subset=["LATITUDE", "LONGITUDE"])

        if map_df.empty:
            st.warning("Data koordinat kosong atau tidak valid.")
        else:
            map_df["Radius"] = map_df["GWh Tahun Ini"].clip(lower=0.05) * 150
            map_df["Warna_Map"] = map_df["Delta GWh"].apply(
                lambda x: [231, 76, 60, 170] if x < 0 else [46, 134, 193, 170]
            )

            if "NAMA PELANGGAN" not in map_df.columns:
                map_df["NAMA PELANGGAN"] = "-"

            view_state = pdk.ViewState(
                latitude=map_df["LATITUDE"].mean(),
                longitude=map_df["LONGITUDE"].mean(),
                zoom=7,
                pitch=0
            )

            scatter_layer = pdk.Layer(
                "ScatterplotLayer",
                data=map_df,
                get_position="[LONGITUDE, LATITUDE]",
                get_radius="Radius",
                get_fill_color="Warna_Map",
                pickable=True,
                auto_highlight=True
            )

            tooltip = {
                "html": """
                <b>{NAMA PELANGGAN}</b><br/>
                UP3: {UP3}<br/>
                Tarif: {TARIF}<br/>
                Cluster: {KLUSTER USAHA}<br/>
                GWh Tahun Ini: {GWh Tahun Ini}<br/>
                Delta GWh: {Delta GWh}<br/>
                Growth: {Growth %}%
                """,
                "style": {
                    "backgroundColor": "white",
                    "color": "black"
                }
            }

            st.pydeck_chart(
                pdk.Deck(
                    layers=[scatter_layer],
                    initial_view_state=view_state,
                    tooltip=tooltip,
                    map_style="light"
                ),
                use_container_width=True
            )

            st.caption("Biru = naik/positif, Merah = turun/negatif. Ukuran bubble mengikuti GWh Tahun Ini.")

# =========================
# GRAFIK GOLONGAN TARIF
# =========================
if "GOLONGAN TARIF" in df_filter.columns:
    st.subheader("📊 Penjualan per Golongan Tarif")

    grafik_golongan = df_filter.groupby(
        "GOLONGAN TARIF", as_index=False
    )[["GWh Tahun Lalu", "GWh Tahun Ini"]].sum()

    fig_golongan = px.bar(
        grafik_golongan,
        x="GOLONGAN TARIF",
        y=["GWh Tahun Lalu", "GWh Tahun Ini"],
        barmode="group",
        text_auto=".2f",
        title=f"Penjualan per Golongan Tarif - {mode_periode} {pilih_bulan}"
    )

    fig_golongan.update_layout(
        title=dict(font=dict(size=22)),
        xaxis=dict(title_font=dict(size=16), tickfont=dict(size=14)),
        yaxis=dict(title_font=dict(size=16), tickfont=dict(size=14)),
        legend=dict(font=dict(size=14))
    )

    st.plotly_chart(fig_golongan, use_container_width=True)

# =========================
# CLUSTER SUMMARY
# =========================
cluster_summary = df_filter.groupby(
    "KLUSTER USAHA", as_index=False
).agg({
    "GWh Tahun Ini": "sum",
    "GWh Tahun Lalu": "sum",
    "Delta GWh": "sum"
})

cluster_summary["Growth %"] = cluster_summary.apply(
    lambda x: (x["Delta GWh"] / x["GWh Tahun Lalu"] * 100)
    if x["GWh Tahun Lalu"] != 0 else 0,
    axis=1
)

top_cluster = cluster_summary.sort_values(metrik_cluster, ascending=False).head(top_n_cluster)
bottom_cluster = cluster_summary.sort_values(metrik_cluster, ascending=True).head(bottom_n_cluster)

if mode_cluster == "Top saja":
    cluster_tampil = top_cluster.copy()
    cluster_tampil["Kategori"] = "Top"
    judul_cluster = f"Top {top_n_cluster} Cluster berdasarkan {metrik_cluster}"

elif mode_cluster == "Bottom saja":
    cluster_tampil = bottom_cluster.copy()
    cluster_tampil["Kategori"] = "Bottom"
    judul_cluster = f"Bottom {bottom_n_cluster} Cluster berdasarkan {metrik_cluster}"

else:
    top_cluster["Kategori"] = "Top"
    bottom_cluster["Kategori"] = "Bottom"
    cluster_tampil = pd.concat([top_cluster, bottom_cluster], ignore_index=True)
    cluster_tampil = cluster_tampil.drop_duplicates(subset=["KLUSTER USAHA"])
    judul_cluster = f"Top {top_n_cluster} dan Bottom {bottom_n_cluster} Cluster berdasarkan {metrik_cluster}"

cluster_tampil = cluster_tampil.sort_values(metrik_cluster, ascending=True)

cluster_tampil["Warna"] = cluster_tampil[metrik_cluster].apply(
    lambda x: "Positif" if x >= 0 else "Negatif"
)

st.subheader("🏭 Top & Bottom Cluster KBLI")

fig_cluster = px.bar(
    cluster_tampil,
    x=metrik_cluster,
    y="KLUSTER USAHA",
    orientation="h",
    text=metrik_cluster,
    color="Warna",
    color_discrete_map={
        "Positif": "#5DADE2",
        "Negatif": "#E74C3C"
    },
    title=f"{judul_cluster} - {mode_periode} {pilih_bulan}"
)

fig_cluster.update_traces(
    texttemplate="%{text:.1f}",
    textposition="inside",
    textfont=dict(size=16, color="white")
)

fig_cluster.update_layout(
    height=max(650, len(cluster_tampil) * 45),
    title=dict(font=dict(size=24)),
    xaxis=dict(
        title=metrik_cluster,
        title_font=dict(size=18),
        tickfont=dict(size=15),
        zeroline=True,
        zerolinewidth=2,
        zerolinecolor="black"
    ),
    yaxis=dict(
        title="KLUSTER USAHA",
        title_font=dict(size=18),
        tickfont=dict(size=15)
    ),
    legend=dict(font=dict(size=15), title_font=dict(size=16))
)

st.plotly_chart(fig_cluster, use_container_width=True)

st.subheader("🍩 Komposisi Penjualan per Cluster")

donut = cluster_summary.sort_values("GWh Tahun Ini", ascending=False).head(15)

fig_donut = px.pie(
    donut,
    names="KLUSTER USAHA",
    values="GWh Tahun Ini",
    hole=0.45,
    title=f"Share Top 15 Cluster - {mode_periode} {pilih_bulan} {tahun_ini}"
)

st.plotly_chart(fig_donut, use_container_width=True)

# =========================
# TABEL PIVOT
# =========================
st.subheader("📋 Tabel Pivot Dinamis")

index_pivot = ["UP3", "TARIF", "KLUSTER USAHA"]

if "GOLONGAN TARIF" in df_filter.columns:
    index_pivot = ["UP3", "TARIF", "GOLONGAN TARIF", "KLUSTER USAHA"]

pivot = df_filter.pivot_table(
    index=index_pivot,
    values=["GWh Tahun Lalu", "GWh Tahun Ini", "Delta GWh"],
    aggfunc="sum"
).reset_index()

pivot["Growth %"] = pivot.apply(
    lambda x: (x["Delta GWh"] / x["GWh Tahun Lalu"] * 100)
    if x["GWh Tahun Lalu"] != 0 else 0,
    axis=1
)

st.dataframe(pivot, use_container_width=True)

# =========================
# AI INSIGHT
# =========================
st.divider()
st.subheader("🧠 Generate AI Insight")

if st.button("🔍 Generate AI Insight"):
    if not api_key:
        st.warning("API Key belum tersedia di Streamlit Secrets.")
    else:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel("gemini-3-flash-preview")

        data_ai_top = pivot.sort_values("Delta GWh", ascending=False).head(15)
        data_ai_bottom = pivot.sort_values("Delta GWh", ascending=True).head(15)
        data_ai = pd.concat([data_ai_top, data_ai_bottom], ignore_index=True)

        prompt = f"""
        Anda adalah Senior Business Analyst PLN UID Jawa Timur.

        KONTEKS:
        - Dashboard Penjualan Tenaga Listrik Tegangan Menengah (TM)
        - Mode periode: {mode_periode}
        - Bulan: {pilih_bulan}
        - Perbandingan: {tahun_lalu} vs {tahun_ini}

        DATA UTAMA:
        {data_ai.to_string(index=False)}

        INSTRUKSI ANALISIS:
        1. Fokus pada perubahan signifikan, baik positif maupun negatif.
        2. Gunakan angka spesifik dalam GWh dan persen.
        3. Identifikasi cluster/customer segment yang menjadi growth driver.
        4. Identifikasi cluster yang turun dan perlu perhatian.
        5. Berikan rekomendasi aksi yang konkret untuk UP3/ULP.
        6. Analisis bisa diselaraskan dengan isu/kondisi yang ada di media secara valid.
        7. Cantumkan sumber data jika mengambil dari eksternal.

        OUTPUT WAJIB:
        1. Executive Summary maksimal 3 kalimat.
        2. Top Growth Drivers.
        3. Underperforming Cluster.
        4. Warning Area.
        5. Rekomendasi Aksi.
        6. Narasi singkat untuk bahan paparan GM UID Jawa Timur.

        GAYA BAHASA:
        Formal, tajam, ringkas, dan cocok untuk bahan paparan manajemen PLN.
        """

        with st.spinner("AI sedang menganalisis data..."):
            response = model.generate_content(prompt)

        st.success("AI Insight berhasil dibuat.")
        st.markdown(response.text)
