import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# =========================
# CLEAN FILE (CHUẨN 100%)
# =========================
def clean_file(file):

    df = pd.read_excel(file, header=None)

    # giữ dòng có MSSV
    df = df[df.apply(lambda row: row.astype(str).str.contains(r"\d{8,}", regex=True).any(), axis=1)]
    df = df.reset_index(drop=True)

    df.columns = range(len(df.columns))

    # tìm MSSV
    mssv_col = df.apply(lambda col: col.astype(str).str.contains(r"\d{8,}", regex=True).sum()).idxmax()

    # ghép họ + tên
    ho = df[mssv_col + 1].astype(str)
    ten = df[mssv_col + 2].astype(str)
    ho_ten = ho + " " + ten

    # =========================
    # tìm cột điểm (numeric thật)
    # =========================
    score_cols = []

    for col in df.columns:
        numeric_ratio = pd.to_numeric(df[col], errors="coerce").notna().mean()
        if numeric_ratio > 0.5:
            score_cols.append(col)

    # chỉ lấy cột sau MSSV
    score_cols = [c for c in score_cols if c > mssv_col][:4]

    # =========================
    # tạo dataframe sạch
    # =========================
    clean = pd.DataFrame()

    clean["MSSV"] = df[mssv_col].astype(str).str.extract(r"(\d+)")
    clean["Họ và Tên"] = ho_ten
    clean["Lớp sinh hoạt"] = df[mssv_col + 3].astype(str)

    clean["Chuyên cần"] = pd.to_numeric(df[score_cols[0]], errors="coerce")
    clean["Giữa kỳ"] = pd.to_numeric(df[score_cols[1]], errors="coerce")
    clean["Thảo luận"] = pd.to_numeric(df[score_cols[2]], errors="coerce")
    clean["Cuối kỳ"] = pd.to_numeric(df[score_cols[3]], errors="coerce")

    # STT
    clean.insert(0, "STT", range(1, len(clean)+1))

    # điểm trung bình
    clean["Điểm trung bình"] = (
        clean["Chuyên cần"].fillna(0)*0.1 +
        clean["Giữa kỳ"].fillna(0)*0.2 +
        clean["Thảo luận"].fillna(0)*0.2 +
        clean["Cuối kỳ"].fillna(0)*0.5
    )

    # xếp loại
    clean["Xếp loại"] = clean["Điểm trung bình"].apply(
        lambda x: "Giỏi" if x >= 8 else
        "Khá" if x >= 6.5 else
        "Trung bình" if x >= 5 else "Yếu"
    )

    clean["Trạng thái"] = clean["Điểm trung bình"].apply(
        lambda x: "Đậu" if x >= 5 else "Rớt"
    )

    clean["Lớp học phần"] = file.name.replace(".xlsx","")

    return clean


# =========================
# LOAD DATA
# =========================
def load_data(files):
    dfs = []
    for f in files:
        try:
            dfs.append(clean_file(f))
        except Exception as e:
            st.error(f"Lỗi {f.name}: {e}")
    return pd.concat(dfs, ignore_index=True)


# =========================
# UI
# =========================
st.set_page_config(layout="wide")
st.title("📊 Dashboard Điểm Toán Cao Cấp")

uploaded_files = st.file_uploader(
    "Upload file Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.warning("⚠️ Upload file trước")
    st.stop()

df = load_data(uploaded_files)

if df.empty:
    st.error("❌ Không có dữ liệu")
    st.stop()

# =========================
# FILTER
# =========================
classes = df["Lớp học phần"].unique()
selected_class = st.sidebar.selectbox("Chọn lớp", classes)

filtered = df[df["Lớp học phần"] == selected_class]

# =========================
# TABS
# =========================
tab1, tab2, tab3 = st.tabs(["📋 Danh sách","📊 Phân tích","📈 Biểu đồ"])

# -------------------------
# TAB 1
# -------------------------
with tab1:
    st.dataframe(filtered, use_container_width=True)

# -------------------------
# TAB 2
# -------------------------
with tab2:
    col1, col2, col3 = st.columns(3)

    col1.metric("Số SV", len(filtered))
    col2.metric("Điểm TB", round(filtered["Điểm trung bình"].mean(),2))
    col3.metric("Cao nhất", round(filtered["Điểm trung bình"].max(),2))

    st.subheader("Xếp loại")
    st.write(filtered["Xếp loại"].value_counts())

    st.subheader("Top 5 sinh viên")
    st.dataframe(filtered.sort_values("Điểm trung bình", ascending=False).head(5))

# -------------------------
# TAB 3
# -------------------------
with tab3:
    vc = filtered["Xếp loại"].value_counts()

    st.subheader("Bar chart")
    st.bar_chart(vc)

    st.subheader("Pie chart")
    fig, ax = plt.subplots()
    vc.plot.pie(autopct='%1.1f%%', ax=ax)
    ax.set_ylabel("")
    st.pyplot(fig)

    st.subheader("Histogram")
    fig2, ax2 = plt.subplots()
    filtered["Điểm trung bình"].plot.hist(bins=10, ax=ax2)
    st.pyplot(fig2)