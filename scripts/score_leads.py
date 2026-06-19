import os
import re
import json
import urllib.request
import pandas as pd

# Google Sheets download URL
SHEET_URL = "https://docs.google.com/spreadsheets/d/16tCAf_qqtgYZxoumYQKMEOdBhKE0wg5A/export?format=xlsx"
LOCAL_EXCEL = "customer_data.xlsx"
LOCAL_JSON = "scored_leads.json"

def download_data():
    """Tải dữ liệu từ Google Sheets"""
    print(f"Downloading data from: {SHEET_URL}")
    urllib.request.urlretrieve(SHEET_URL, LOCAL_EXCEL)
    print("Data downloaded successfully!")

def score_text(text):
    """
    Chấm điểm khách hàng tiềm năng dựa trên tieu_chi_cham_diem.txt.
    Base score: 100. Range: 0 đến 200.
    """
    if not isinstance(text, str):
        return 100, ["Không có thông tin mô tả nhu cầu"], "Neutral"
    
    normalized_text = text.lower()
    score = 100
    reasons = []
    
    # -------------------------------------------------------------
    # 1. TIÊU CHÍ CỘNG 50 ĐIỂM (VIP)
    # -------------------------------------------------------------
    vip_reasons = []
    
    # Ngân sách lớn: Số tiền >= 20 tỷ hoặc cụm từ cụ thể
    budget_match = re.search(r'(\d+(?:\.\d+)?)\s*(?:tỷ|ty|tỉ)', normalized_text)
    has_large_budget = False
    if budget_match:
        val = float(budget_match.group(1).replace(",", "."))
        if val >= 20.0:
            has_large_budget = True
            vip_reasons.append(f"Ngân sách lớn ({int(val)} tỷ)")
    if "tài chính mạnh" in normalized_text or "không thành vấn đề" in normalized_text:
        has_large_budget = True
        vip_reasons.append("Tài chính rất mạnh")
        
    # Loại hình cao cấp
    premium_keywords = ["biệt thự", "penthouse", "shophouse", "quỹ đất công nghiệp", "đất công nghiệp", "sàn văn phòng"]
    matched_premium = [k for k in premium_keywords if k in normalized_text]
    if matched_premium:
        vip_reasons.append(f"Loại hình cao cấp ({', '.join([k.capitalize() for k in matched_premium])})")
        
    # Vị trí đắc địa
    locations = ["quận 1", "ven sông", "vinhomes ocean park", "phú mỹ hưng"]
    matched_locations = [l for l in locations if l in normalized_text]
    if matched_locations:
        vip_reasons.append(f"Vị trí đắc địa ({', '.join([l.title() for l in matched_locations])})")
        
    # Đối tượng khách hàng VIP
    target_clients = ["chủ doanh nghiệp", "nhà đầu tư chuyên nghiệp", "mua sỉ", "mua số lượng lớn"]
    matched_targets = [t for t in target_clients if t in normalized_text]
    if matched_targets:
        vip_reasons.append(f"Khách hàng VIP ({', '.join([t.capitalize() for t in matched_targets])})")
        
    # Tính cấp thiết & Minh bạch
    urgency_keywords = ["pháp lý chuẩn", "sổ hồng riêng", "gặp trực tiếp chủ đầu tư", "pháp lý sạch"]
    matched_urgency = [u for u in urgency_keywords if u in normalized_text]
    if matched_urgency:
        vip_reasons.append(f"Yêu cầu pháp lý & Minh bạch ({', '.join([u.capitalize() for u in matched_urgency])})")
        
    # Áp dụng điểm cộng nếu có tiêu chí VIP (tối đa +100 điểm, tương đương score = 200)
    if vip_reasons:
        score += 50 * len(vip_reasons)
        reasons.extend(vip_reasons)
        
    # -------------------------------------------------------------
    # 2. TIÊU CHÍ TRỪ 50 ĐIỂM (JUNK)
    # -------------------------------------------------------------
    junk_reasons = []
    
    # Yêu cầu phi thực tế (VD: Quận 1 hoặc trung tâm có giá rẻ 1-2 tỷ, trăm triệu)
    is_unrealistic = False
    if ("quận 1" in normalized_text or "trung tâm" in normalized_text) and \
       any(p in normalized_text for p in ["1 tỷ", "2 tỷ", "trăm triệu", "vài trăm", "500 triệu", "dưới 2 tỷ"]):
        is_unrealistic = True
        junk_reasons.append("Yêu cầu phi thực tế (Giá quá thấp tại Quận 1/Trung tâm)")
        
    # Không có nhu cầu
    no_need = ["nhầm số", "không có nhu cầu", "dữ liệu cũ", "nhầm ngành"]
    matched_no_need = [n for n in no_need if n in normalized_text]
    if matched_no_need:
        junk_reasons.append(f"Không có nhu cầu ({', '.join([n.capitalize() for n in matched_no_need])})")
        
    # Khách hàng không thiện chí
    no_goodwill = ["hỏi giá cho vui", "cho vui", "chưa có ý định mua", "thái độ không hợp tác", "không hợp tác"]
    matched_goodwill = [g for g in no_goodwill if g in normalized_text]
    if matched_goodwill:
        junk_reasons.append("Khách hàng thiếu thiện chí")
        
    # Spam/Quảng cáo
    spam_keywords = ["bảo hiểm", "vay vốn", "mời chào", "quảng cáo dịch vụ"]
    matched_spam = [s for s in spam_keywords if s in normalized_text]
    if matched_spam:
        junk_reasons.append(f"Spam/Quảng cáo ({', '.join([s.capitalize() for s in matched_spam])})")
        
    # Thông tin liên lạc lỗi
    comm_error = ["thuê bao", "không bắt máy", "gọi nhiều lần không", "không phản hồi zalo"]
    matched_comm = [c for c in comm_error if c in normalized_text]
    if matched_comm:
        junk_reasons.append("Lỗi liên lạc/Không nghe máy/Không phản hồi")
        
    # Áp dụng điểm trừ nếu có tiêu chí Junk
    if junk_reasons:
        score -= 50 * len(junk_reasons)
        reasons.extend(junk_reasons)
        
    # Giới hạn điểm số từ 0 đến 200
    score = max(0, min(200, score))
    
    # Xác định trạng thái dựa trên điểm số
    if score >= 150:
        status = "VIP"
    elif score <= 50:
        status = "Junk"
    else:
        status = "Neutral"
        
    if not reasons:
        reasons = ["Nhu cầu cơ bản / Đang cân nhắc thêm"]
        
    return score, reasons, status

def process_and_save():
    """Đọc file Excel, chấm điểm và lưu trữ bảo toàn các thay đổi thủ công của con người"""
    if not os.path.exists(LOCAL_EXCEL):
        download_data()
        
    df = pd.read_excel(LOCAL_EXCEL)
    df = df.fillna("")
    
    # Tải dữ liệu cũ nếu đã tồn tại để giữ các chỉnh sửa thủ công (Human-in-the-loop)
    existing_leads = {}
    if os.path.exists(LOCAL_JSON):
        try:
            with open(LOCAL_JSON, "r", encoding="utf-8") as f:
                old_data = json.load(f)
                for item in old_data:
                    existing_leads[item["id"]] = item
        except Exception as e:
            print(f"Could not read old JSON: {e}. Overwriting.")
            
    scored_list = []
    
    for idx, row in df.iterrows():
        lead_id = int(row["id"])
        ten_khach = str(row["ten_khach"])
        sdt = str(row["sdt"])
        nhu_cau_mo_ta = str(row["nhu_cau_mo_ta"])
        
        # Nếu đã có trong cơ sở dữ liệu và được đánh dấu chỉnh sửa thủ công hoặc phê duyệt
        if lead_id in existing_leads and existing_leads[lead_id].get("manual_override", False):
            # Giữ nguyên bản ghi cũ đã chỉnh sửa bởi con người
            scored_list.append(existing_leads[lead_id])
        else:
            score, reasons, status = score_text(nhu_cau_mo_ta)
            
            # Format số điện thoại hiển thị đẹp (thêm số 0 ở đầu nếu chưa có)
            sdt_display = sdt.split(".")[0]
            if sdt_display and not sdt_display.startswith("0") and len(sdt_display) >= 9:
                sdt_display = "0" + sdt_display
                
            lead_record = {
                "id": lead_id,
                "ten_khach": ten_khach,
                "sdt": sdt_display,
                "nhu_cau_mo_ta": nhu_cau_mo_ta,
                "score": score,
                "reasons": reasons,
                "status": status,
                "manual_override": False,
                "comment": ""
            }
            scored_list.append(lead_record)
            
    # Lưu vào JSON file dạng UTF-8
    with open(LOCAL_JSON, "w", encoding="utf-8") as f:
        json.dump(scored_list, f, ensure_ascii=False, indent=2)
        
    print(f"Processed and saved {len(scored_list)} clients to {LOCAL_JSON}!")

if __name__ == "__main__":
    process_and_save()
