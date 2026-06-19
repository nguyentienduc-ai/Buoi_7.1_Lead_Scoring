import streamlit as st
import json
import os
import pandas as pd
import subprocess

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="AI Lead Scoring & Automation System",
    page_icon="🏢",
    layout="wide"
)

LOCAL_JSON = "scored_leads.json"
LOCAL_EXCEL = "customer_data.xlsx"

# Helper: load data
def load_data():
    if not os.path.exists(LOCAL_JSON):
        run_sync()
    with open(LOCAL_JSON, "r", encoding="utf-8") as f:
        return json.load(f)

# Helper: save data
def save_data(data):
    with open(LOCAL_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# Helper: trigger sync script
def run_sync():
    from scripts.score_leads import process_and_save
    process_and_save()

# -------------------------------------------------------------
# DỰNG GIAO DIỆN STREAMLIT
# -------------------------------------------------------------
st.title("🏢 AI Lead Scoring & Automation System")
st.markdown("Hệ thống tự động chấm điểm khách hàng tiềm năng Bất động sản (Human-in-the-loop)")

# Sidebar các nút chức năng
with st.sidebar:
    st.header("⚡ Thao tác hệ thống")
    if st.button("🔄 Đồng bộ dữ liệu từ Sheets", use_container_width=True):
        with st.spinner("Đang tải dữ liệu và chấm điểm..."):
            run_sync()
            st.success("Đồng bộ dữ liệu thành công!")
            
    # Xuất Excel
    st.markdown("---")
    st.header("📤 Xuất báo cáo")
    if os.path.exists(LOCAL_JSON):
        # Chạy server.py để tạo file excel hoặc tạo trực tiếp bằng code
        from server import generate_excel_export
        if st.button("Generate Excel Report", use_container_width=True):
            generate_excel_export()
            st.success("Đã tạo xong file scored_leads_final.xlsx!")
            
        # Thêm nút Download file Excel
        if os.path.exists("scored_leads_final.xlsx"):
            with open("scored_leads_final.xlsx", "rb") as f:
                st.download_button(
                    label="📥 Tải file Excel (.xlsx)",
                    data=f,
                    file_name="scored_leads_final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

# Load data vào state
try:
    leads = load_data()
except Exception as e:
    st.error(f"Không thể đọc cơ sở dữ liệu: {e}")
    leads = []

if leads:
    # -------------------------------------------------------------
    # KPI METRICS KHU VỰC TRÊN CÙNG
    # -------------------------------------------------------------
    total_leads = len(leads)
    vip_leads = sum(1 for l in leads if l.get("status") == "VIP")
    neutral_leads = sum(1 for l in leads if l.get("status") == "Neutral")
    junk_leads = sum(1 for l in leads if l.get("status") == "Junk")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Tổng khách hàng", total_leads)
    col2.metric("Khách hàng VIP", vip_leads, delta=None, delta_color="normal")
    col3.metric("Khách hàng Trung bình", neutral_leads)
    col4.metric("Khách hàng Rác / Spam", junk_leads)

    st.markdown("---")

    # -------------------------------------------------------------
    # BỘ LỌC & TÌM KIẾM
    # -------------------------------------------------------------
    st.subheader("🔍 Danh sách Khách hàng cần kiểm duyệt")
    
    col_s, col_f = st.columns([2, 1])
    search_query = col_s.text_input("Tìm kiếm theo tên, sđt hoặc mô tả nhu cầu", "")
    status_filter = col_f.selectbox("Bộ lọc trạng thái", ["Tất cả", "VIP", "Neutral", "Junk"])

    # Lọc dữ liệu
    filtered_leads = leads
    if status_filter != "Tất cả":
        filtered_leads = [l for l in filtered_leads if l.get("status") == status_filter]
        
    if search_query:
        q = search_query.lower()
        filtered_leads = [
            l for l in filtered_leads 
            if q in l.get("ten_khach", "").lower() 
            or q in l.get("sdt", "") 
            or q in l.get("nhu_cau_mo_ta", "").lower()
        ]

    # Hiển thị bảng dạng DataFrame để xem nhanh
    df_display = pd.DataFrame([
        {
            "ID": l["id"],
            "Họ Tên": l["ten_khach"],
            "SĐT": l["sdt"],
            "Mô tả Nhu Cầu": l["nhu_cau_mo_ta"],
            "Điểm AI": l["score"],
            "Phân Loại": l["status"],
            "Duyệt Thủ Công": "Đã duyệt" if l["manual_override"] else "Chưa duyệt",
            "Ghi chú kiểm duyệt": l["comment"]
        }
        for l in filtered_leads
    ])
    # Render bảng bằng Markdown để tránh import pyarrow (lỗi trên môi trường ARM64 Windows)
    markdown_table = "| ID | Họ Tên | SĐT | Mô tả Nhu Cầu | Điểm AI | Phân Loại | Duyệt | Ghi chú |\n"
    markdown_table += "| :---: | :--- | :---: | :--- | :---: | :---: | :---: | :--- |\n"
    for _, row in df_display.iterrows():
        desc = str(row["Mô tả Nhu Cầu"]).replace("|", "\\|").replace("\n", " ")
        comm = str(row["Ghi chú kiểm duyệt"]).replace("|", "\\|").replace("\n", " ")
        markdown_table += f"| {row['ID']} | {row['Họ Tên']} | {row['SĐT']} | {desc} | **{row['Điểm AI']}** | `{row['Phân Loại']}` | {row['Duyệt Thủ Công']} | *{comm}* |\n"
    
    st.markdown(markdown_table, unsafe_allow_html=True)

    # -------------------------------------------------------------
    # KHU VỰC HUMAN-IN-THE-LOOP (CHỈNH SỬA / DUYỆT LEAD)
    # -------------------------------------------------------------
    st.markdown("---")
    st.subheader("✍️ Kiểm duyệt & Cập nhật trạng thái khách hàng")
    
    selected_id = st.selectbox("Chọn ID Khách hàng muốn kiểm duyệt:", [l["id"] for l in filtered_leads])
    
    if selected_id:
        # Lấy thông tin lead được chọn
        lead = next((l for l in leads if l["id"] == selected_id), None)
        if lead:
            st.info(f"**Khách hàng:** {lead['ten_khach']} | **Số điện thoại:** {lead['sdt']}\n\n**Nhu cầu mô tả:** {lead['nhu_cau_mo_ta']}")
            
            with st.form("edit_form"):
                col_score, col_status = st.columns(2)
                
                # Input Điểm số
                new_score = col_score.number_input(
                    "Điểm số tiềm năng (0 - 200)", 
                    min_value=0, 
                    max_value=200, 
                    value=int(lead["score"])
                )
                
                # Phân loại trạng thái
                status_options = ["VIP", "Neutral", "Junk"]
                default_status_idx = status_options.index(lead["status"]) if lead["status"] in status_options else 1
                new_status = col_status.selectbox("Trạng thái", status_options, index=default_status_idx)
                
                # Ghi chú kiểm duyệt
                new_comment = st.text_area("Ghi chú kiểm duyệt / Chỉ dẫn bán hàng", value=lead.get("comment", ""))
                
                submitted = st.form_submit_button("💾 Xác nhận & Lưu kết quả duyệt")
                
                if submitted:
                    # Cập nhật thông tin
                    for l in leads:
                        if l["id"] == selected_id:
                            l["score"] = int(new_score)
                            l["status"] = new_status
                            l["comment"] = new_comment
                            l["manual_override"] = True
                            break
                    save_data(leads)
                    st.success(f"Đã cập nhật trạng thái khách hàng ID {selected_id} thành công!")
                    st.rerun()
else:
    st.warning("Hiện tại chưa có dữ liệu khách hàng. Vui lòng bấm nút 'Đồng bộ dữ liệu từ Sheets' ở thanh bên trái.")
