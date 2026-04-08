import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# ======================================================
# CONFIG
# ======================================================
st.set_page_config(
    page_title="Dashboard Điểm Toán Cao Cấp",
    page_icon="📊",
    layout="wide"
)

# ======================================================
# COLOR PALETTE
# ======================================================
CLASS_COLORS = [
    "#2563EB",  # blue
    "#7C3AED",  # violet
    "#059669",  # emerald
    "#EA580C",  # orange
    "#DC2626",  # red
    "#0891B2"   # cyan
]

RANK_COLOR_MAP = {
    "Giỏi": "#2563EB",
    "Khá": "#059669",
    "Trung bình": "#F59E0B",
    "Yếu": "#DC2626"
}

STATUS_COLOR_MAP = {
    "Đậu": "#16A34A",
    "Rớt": "#DC2626"
}

# ======================================================
# STYLE
# ======================================================
st.markdown("""
<style>
:root {
    --bg: #f8fafc;
    --card: #ffffff;
    --border: #e5e7eb;
    --text: #0f172a;
    --muted: #64748b;
    --primary: #2563eb;
}

.stApp {
    background: linear-gradient(180deg, #f8fbff 0%, #f8fafc 100%);
}

.block-container {
    padding-top: 1.2rem;
    padding-bottom: 1rem;
    max-width: 96%;
}

h1, h2, h3 {
    color: var(--text);
    font-weight: 700;
}

div[data-testid="metric-container"] {
    background: linear-gradient(135deg, #ffffff 0%, #f8fbff 100%);
    border: 1px solid var(--border);
    padding: 16px 18px;
    border-radius: 18px;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
}

[data-testid="stMetricLabel"] {
    color: #64748b;
    font-weight: 600;
}

[data-testid="stMetricValue"] {
    color: #0f172a;
    font-size: 30px;
    font-weight: 800;
}

div.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
}

div.stTabs [data-baseweb="tab"] {
    height: 46px;
    border-radius: 12px 12px 0 0;
    padding: 0 18px;
    background: #f1f5f9;
    color: #334155;
    font-weight: 600;
}

div.stTabs [aria-selected="true"] {
    background: #ffffff !important;
    color: #2563eb !important;
    border-top: 3px solid #2563eb !important;
}

section[data-testid="stSidebar"] {
    background: #ffffff;
    border-right: 1px solid #e5e7eb;
}

.small-note {
    color: #64748b;
    font-size: 14px;
    margin-top: -4px;
}

.title-box {
    padding: 18px 20px;
    border-radius: 18px;
    background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
    color: white;
    margin-bottom: 12px;
    box-shadow: 0 10px 30px rgba(37, 99, 235, 0.18);
}

.title-box h1 {
    color: white !important;
    margin: 0;
    font-size: 2.3rem;
}

.title-box p {
    margin: 6px 0 0 0;
    color: #e0ecff;
    font-size: 15px;
}
</style>
""", unsafe_allow_html=True)

# ======================================================
# HELPERS
# ======================================================
def apply_chart_style(fig, height=430):
    fig.update_layout(
        height=height,
        paper_bgcolor="white",
        plot_bgcolor="white",
        margin=dict(l=30, r=30, t=60, b=30),
        font=dict(color="#0f172a", size=13),
        title_font=dict(size=18, color="#0f172a"),
        legend_title_text="",
    )
    fig.update_xaxes(
        showgrid=False,
        linecolor="#CBD5E1",
        tickfont=dict(color="#334155")
    )
    fig.update_yaxes(
        gridcolor="#E2E8F0",
        zerolinecolor="#E2E8F0",
        tickfont=dict(color="#334155")
    )
    return fig

def find_header_row(df_raw):
    for i in range(min(15, len(df_raw))):
        row_text = " ".join(df_raw.iloc[i].fillna("").astype(str).tolist()).lower()
        if "mã số sinh viên" in row_text and "điểm tổng hợp" in row_text:
            return i
    raise ValueError("Không tìm thấy dòng tiêu đề dữ liệu.")

def clean_file(file):
    df_raw = pd.read_excel(file, header=None)
    header_row = find_header_row(df_raw)

    h1 = df_raw.iloc[header_row].fillna("")
    h2 = df_raw.iloc[header_row + 1].fillna("")

    merged_headers = []
    for a, b in zip(h1, h2):
        a = str(a).strip()
        b = str(b).strip()
        merged = f"{a} {b}".strip()
        merged = " ".join(merged.split())
        merged_headers.append(merged)

    df = df_raw.iloc[header_row + 2:].copy().reset_index(drop=True)
    df.columns = merged_headers
    df = df.dropna(how="all").copy()

    # xử lý tên cột trùng nhau
    def get_series_by_col(col_name):
        s = df[col_name]
        if isinstance(s, pd.DataFrame):
            s = s.iloc[:, 0]
        return s

    # tìm cột MSSV
    mssv_col = None
    for col in df.columns:
        col_name = str(col).lower()
        col_data = get_series_by_col(col).astype(str)
        if "mã số sinh viên" in col_name and col_data.str.contains(r"\d{8,}", regex=True, na=False).sum() > 0:
            mssv_col = col
            break

    if mssv_col is None:
        for col in df.columns:
            col_data = get_series_by_col(col).astype(str)
            if col_data.str.contains(r"\d{8,}", regex=True, na=False).sum() > 0:
                mssv_col = col
                break

    if mssv_col is None:
        raise ValueError("Không tìm được cột MSSV.")

    def find_col(keyword_list):
        for col in df.columns:
            txt = str(col).lower()
            if all(k in txt for k in keyword_list):
                return col
        return None

    lopsh_col = find_col(["lớp", "sinh", "hoạt"])
    tong_col = find_col(["điểm", "tổng", "hợp"])

    if not lopsh_col:
        raise ValueError("Không tìm được cột Lớp sinh hoạt.")
    if not tong_col:
        raise ValueError("Không tìm được cột Điểm tổng hợp.")

    # chỉ giữ dòng sinh viên thật
    mssv_series = get_series_by_col(mssv_col).astype(str)
    df = df[mssv_series.str.contains(r"\d{8,}", regex=True, na=False)].copy().reset_index(drop=True)

    # helper sau khi đã filter
    def get_series_filtered(col_name):
        s = df[col_name]
        if isinstance(s, pd.DataFrame):
            s = s.iloc[:, 0]
        return s

    # MSSV: giữ nguyên chuỗi số, không lstrip số 0
    df["MSSV"] = (
        get_series_filtered(mssv_col)
        .astype(str)
        .str.strip()
        .str.replace(".0","",regex=False)
        .str.zfill(8)
    )

    # GHÉP HỌ VÀ TÊN:
    # lấy tất cả cột nằm giữa MSSV và Lớp sinh hoạt
    col_list = list(df.columns)
    mssv_idx = col_list.index(mssv_col)
    lop_idx = col_list.index(lopsh_col)

    if lop_idx <= mssv_idx + 1:
        df["Họ và Tên"] = ""
    else:
        name_values = df.iloc[:, mssv_idx + 1:lop_idx].copy()

        for i in range(name_values.shape[1]):
            name_values.iloc[:, i] = (
                name_values.iloc[:, i]
                .fillna("")
                .astype(str)
                .replace("nan", "")
                .replace("None", "")
                .replace("none", "")
                .str.strip()
            )

        df["Họ và Tên"] = (
            name_values.apply(
                lambda row: " ".join([x for x in row.tolist() if str(x).strip() != ""]),
                axis=1
            )
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

    df["Lớp sinh hoạt"] = get_series_filtered(lopsh_col).astype(str).str.strip()

    # Lấy 4 cột điểm thành phần bằng vị trí:
    # các cột số nằm giữa Lớp sinh hoạt và Điểm tổng hợp
    tong_idx = col_list.index(tong_col)
    middle_block = df.iloc[:, lop_idx + 1:tong_idx].copy()

    numeric_cols = []
    for i in range(middle_block.shape[1]):
        col_data = pd.to_numeric(middle_block.iloc[:, i], errors="coerce")
        if col_data.notna().mean() > 0.5:
            numeric_cols.append(middle_block.columns[i])

    # lấy đúng 4 cột điểm thành phần cuối cùng
    if len(numeric_cols) < 4:
        raise ValueError("Không tìm đủ 4 cột điểm thành phần.")
    score_cols = numeric_cols[-4:]

    df["Chuyên cần"] = pd.to_numeric(df[score_cols[0]], errors="coerce")
    df["Giữa kỳ"] = pd.to_numeric(df[score_cols[1]], errors="coerce")
    df["Thảo luận"] = pd.to_numeric(df[score_cols[2]], errors="coerce")
    df["Cuối kỳ"] = pd.to_numeric(df[score_cols[3]], errors="coerce")

    tong_hop = pd.to_numeric(get_series_filtered(tong_col), errors="coerce")

    clean = pd.DataFrame({
        "STT": range(1, len(df) + 1),
        "MSSV": df["MSSV"],
        "Họ và Tên": df["Họ và Tên"],
        "Lớp sinh hoạt": df["Lớp sinh hoạt"],
        "Chuyên cần": df["Chuyên cần"],
        "Giữa kỳ": df["Giữa kỳ"],
        "Thảo luận": df["Thảo luận"],
        "Cuối kỳ": df["Cuối kỳ"],
        "Điểm trung bình": tong_hop.round(2)
    })

    clean["Xếp loại"] = clean["Điểm trung bình"].apply(
        lambda x: "Giỏi" if x >= 8 else
        "Khá" if x >= 6.5 else
        "Trung bình" if x >= 5 else "Yếu"
    )

    clean["Trạng thái"] = clean["Điểm trung bình"].apply(
        lambda x: "Đậu" if x >= 5 else "Rớt"
    )

    clean["Lớp học phần"] = file.name.replace(".xlsx", "")
    return clean

@st.cache_data
def load_data(files):
    dfs = []
    errors = []
    for f in files:
        try:
            dfs.append(clean_file(f))
        except Exception as e:
            errors.append(f"{f.name}: {e}")
    return dfs, errors

# ======================================================
# HEADER
# ======================================================
st.markdown("""
<div class="title-box">
    <h1>📊 Dashboard Điểm Toán Cao Cấp</h1>
    <p>Phân tích, so sánh và trực quan hóa kết quả học tập giữa các lớp học phần</p>
</div>
""", unsafe_allow_html=True)

# ======================================================
# UPLOAD
# ======================================================
uploaded_files = st.file_uploader(
    "Upload file Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("Hãy upload các file Excel của các lớp để bắt đầu.")
    st.stop()

dfs, errors = load_data(uploaded_files)

if errors:
    for err in errors:
        st.error(err)

if not dfs:
    st.error("Không có dữ liệu hợp lệ.")
    st.stop()

df = pd.concat(dfs, ignore_index=True)

# ======================================================
# SIDEBAR
# ======================================================
st.sidebar.header("⚙️ Bộ lọc phân tích")

all_classes = sorted(df["Lớp học phần"].dropna().unique().tolist())

selected_classes = st.sidebar.multiselect(
    "Chọn lớp để so sánh",
    options=all_classes,
    default=all_classes
)

if not selected_classes:
    st.warning("Vui lòng chọn ít nhất 1 lớp.")
    st.stop()

filtered = df[df["Lớp học phần"].isin(selected_classes)].copy()

# map màu theo lớp
class_color_map = {
    cls: CLASS_COLORS[i % len(CLASS_COLORS)]
    for i, cls in enumerate(all_classes)
}

# ======================================================
# KPI
# ======================================================
total_students = len(filtered)
avg_score = filtered["Điểm trung bình"].mean()
max_score = filtered["Điểm trung bình"].max()
min_score = filtered["Điểm trung bình"].min()
pass_rate = filtered["Trạng thái"].eq("Đậu").mean() * 100

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Tổng số SV", f"{total_students}")
k2.metric("Điểm TB", f"{avg_score:.2f}")
k3.metric("Cao nhất", f"{max_score:.2f}")
k4.metric("Thấp nhất", f"{min_score:.2f}")
k5.metric("Tỉ lệ đậu", f"{pass_rate:.1f}%")

# ======================================================
# TABS
# ======================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Danh sách dữ liệu",
    "📊 So sánh các lớp",
    "📈 Phân phối & tương quan",
    "🔍 Chi tiết 1 lớp"
])

# ======================================================
# TAB 1
# ======================================================
with tab1:
    st.subheader("Danh sách sinh viên")
    st.caption("Hiển thị dữ liệu của các lớp đang được chọn")

    st.dataframe(filtered, width="stretch", hide_index=True)

    st.subheader("Tổng hợp theo lớp")
    summary = (
        filtered.groupby("Lớp học phần")
        .agg(
            Số_SV=("MSSV", "count"),
            Điểm_TB=("Điểm trung bình", "mean"),
            Cao_nhất=("Điểm trung bình", "max"),
            Thấp_nhất=("Điểm trung bình", "min"),
            Tỉ_lệ_đậu=("Trạng thái", lambda x: x.eq("Đậu").mean() * 100)
        )
        .reset_index()
        .round(2)
    )

    st.dataframe(summary, width="stretch", hide_index=True)

# ======================================================
# TAB 2
# ======================================================
with tab2:
    st.subheader("So sánh kết quả giữa các lớp")

    compare_df = (
        filtered.groupby("Lớp học phần")
        .agg(
            Số_SV=("MSSV", "count"),
            Điểm_TB=("Điểm trung bình", "mean"),
            Chuyên_cần=("Chuyên cần", "mean"),
            Giữa_kỳ=("Giữa kỳ", "mean"),
            Thảo_luận=("Thảo luận", "mean"),
            Cuối_kỳ=("Cuối kỳ", "mean")
        )
        .reset_index()
        .round(2)
    )

    c1, c2 = st.columns(2)

    with c1:
        fig_bar = px.bar(
            compare_df,
            x="Lớp học phần",
            y="Điểm_TB",
            text="Điểm_TB",
            title="Điểm trung bình từng lớp",
            color="Lớp học phần",
            color_discrete_map=class_color_map
        )
        fig_bar.update_traces(textposition="outside")
        st.plotly_chart(apply_chart_style(fig_bar, 430), width="stretch")

    with c2:
        long_score = compare_df.melt(
            id_vars="Lớp học phần",
            value_vars=["Chuyên_cần", "Giữa_kỳ", "Thảo_luận", "Cuối_kỳ"],
            var_name="Thành phần",
            value_name="Điểm TB"
        )

        fig_line = px.line(
            long_score,
            x="Thành phần",
            y="Điểm TB",
            color="Lớp học phần",
            markers=True,
            title="Xu hướng điểm thành phần giữa các lớp",
            color_discrete_map=class_color_map
        )
        st.plotly_chart(apply_chart_style(fig_line, 430), width="stretch")

    st.subheader("Cơ cấu xếp loại theo phần trăm")

    rank_percent = (
        filtered.groupby(["Lớp học phần", "Xếp loại"])
        .size()
        .reset_index(name="Số lượng")
    )
    total_each = rank_percent.groupby("Lớp học phần")["Số lượng"].transform("sum")
    rank_percent["Tỉ lệ %"] = (rank_percent["Số lượng"] / total_each * 100).round(1)

    fig_stack = px.bar(
        rank_percent,
        x="Lớp học phần",
        y="Tỉ lệ %",
        color="Xếp loại",
        text="Tỉ lệ %",
        title="Tỉ trọng xếp loại của từng lớp (%)",
        color_discrete_map=RANK_COLOR_MAP
    )
    fig_stack.update_layout(barmode="stack")
    st.plotly_chart(apply_chart_style(fig_stack, 470), width="stretch")

    c3, c4 = st.columns(2)

    with c3:
        rank_count = (
            filtered.groupby(["Lớp học phần", "Xếp loại"])
            .size()
            .reset_index(name="Số lượng")
        )

        fig_group = px.bar(
            rank_count,
            x="Lớp học phần",
            y="Số lượng",
            color="Xếp loại",
            barmode="group",
            text="Số lượng",
            title="Số lượng Giỏi / Khá / Trung bình / Yếu",
            color_discrete_map=RANK_COLOR_MAP
        )
        st.plotly_chart(apply_chart_style(fig_group, 430), width="stretch")

    with c4:
        pass_df = (
            filtered.groupby(["Lớp học phần", "Trạng thái"])
            .size()
            .reset_index(name="Số lượng")
        )

        fig_pass = px.bar(
            pass_df,
            x="Lớp học phần",
            y="Số lượng",
            color="Trạng thái",
            barmode="stack",
            text="Số lượng",
            title="Cơ cấu Đậu / Rớt theo lớp",
            color_discrete_map=STATUS_COLOR_MAP
        )
        st.plotly_chart(apply_chart_style(fig_pass, 430), width="stretch")

# ======================================================
# TAB 3
# ======================================================
with tab3:
    st.subheader("Phân phối điểm và tương quan giữa các biến")

    p1, p2 = st.columns(2)

    with p1:
        fig_hist = px.histogram(
            filtered,
            x="Điểm trung bình",
            color="Lớp học phần",
            nbins=15,
            barmode="overlay",
            opacity=0.65,
            title="Histogram - Phân phối điểm trung bình",
            color_discrete_map=class_color_map
        )
        st.plotly_chart(apply_chart_style(fig_hist, 430), width="stretch")

    with p2:
        fig_box = px.box(
            filtered,
            x="Lớp học phần",
            y="Điểm trung bình",
            color="Lớp học phần",
            points="outliers",
            title="Boxplot - Phân phối điểm theo lớp",
            color_discrete_map=class_color_map
        )
        st.plotly_chart(apply_chart_style(fig_box, 430), width="stretch")

    p3, p4 = st.columns(2)

    with p3:
        fig_scatter = px.scatter(
        filtered,
        x="Cuối kỳ",
        y="Điểm trung bình",
        color="Lớp học phần",
        size="Điểm trung bình",
        hover_data=["MSSV", "Họ và Tên", "Giữa kỳ", "Cuối kỳ", "Xếp loại"],
        title="Scatter - Quan hệ giữa điểm cuối kỳ và điểm trung bình",
        color_discrete_map=class_color_map
        )
        st.plotly_chart(apply_chart_style(fig_scatter, 470), width="stretch")

    with p4:
        corr_cols = ["Chuyên cần", "Giữa kỳ", "Thảo luận", "Cuối kỳ", "Điểm trung bình"]
        corr = filtered[corr_cols].corr().round(2)

        fig_heat = px.imshow(
            corr,
            text_auto=True,
            aspect="auto",
            title="Heatmap - Tương quan giữa các biến số",
            color_continuous_scale="Blues"
        )
        fig_heat.update_xaxes(side="bottom")
        st.plotly_chart(apply_chart_style(fig_heat, 470), width="stretch")

# ======================================================
# TAB 4
# ======================================================
with tab4:
    st.subheader("Chi tiết từng lớp")

    detail_class = st.selectbox(
        "Chọn lớp để xem chi tiết",
        options=selected_classes,
        key="detail_class_tab"
    )

    detail_df = filtered[filtered["Lớp học phần"] == detail_class].copy()

    d1, d2, d3, d4 = st.columns(4)
    d1.metric("Số sinh viên", len(detail_df))
    d2.metric("Điểm TB lớp", f"{detail_df['Điểm trung bình'].mean():.2f}")
    d3.metric("Điểm cao nhất", f"{detail_df['Điểm trung bình'].max():.2f}")
    d4.metric("Tỉ lệ đậu", f"{detail_df['Trạng thái'].eq('Đậu').mean() * 100:.1f}%")

    left, right = st.columns([1.3, 1])

    with left:
        st.markdown("### Top 10 sinh viên")
        top10 = detail_df.sort_values("Điểm trung bình", ascending=False).head(10)
        st.dataframe(top10, width="stretch", hide_index=True)

    with right:
        pie_df = detail_df["Xếp loại"].value_counts().reset_index()
        pie_df.columns = ["Xếp loại", "Số lượng"]

        fig_pie = px.pie(
            pie_df,
            names="Xếp loại",
            values="Số lượng",
            hole=0.5,
            title=f"Tỉ lệ xếp loại - {detail_class}",
            color="Xếp loại",
            color_discrete_map=RANK_COLOR_MAP
        )
        st.plotly_chart(apply_chart_style(fig_pie, 400), width="stretch")

    b1, b2 = st.columns(2)

    with b1:
        score_avg = pd.DataFrame({
            "Thành phần": ["Chuyên cần", "Giữa kỳ", "Thảo luận", "Cuối kỳ"],
            "Điểm TB": [
                detail_df["Chuyên cần"].mean(),
                detail_df["Giữa kỳ"].mean(),
                detail_df["Thảo luận"].mean(),
                detail_df["Cuối kỳ"].mean()
            ]
        }).round(2)

        fig_component = px.bar(
            score_avg,
            x="Thành phần",
            y="Điểm TB",
            text="Điểm TB",
            title="Điểm trung bình từng thành phần"
        )
        fig_component.update_traces(
            textposition="outside",
            marker_color="#2563EB"
        )
        st.plotly_chart(apply_chart_style(fig_component, 400), width="stretch")

    with b2:
        fig_hist_detail = px.histogram(
            detail_df,
            x="Điểm trung bình",
            nbins=12,
            title=f"Phân phối điểm - {detail_class}"
        )
        fig_hist_detail.update_traces(marker_color="#7C3AED")
        st.plotly_chart(apply_chart_style(fig_hist_detail, 400), width="stretch")

    st.markdown("### Danh sách đầy đủ của lớp")
    st.dataframe(
        detail_df.sort_values("Điểm trung bình", ascending=False),
        width="stretch",
        hide_index=True
    )