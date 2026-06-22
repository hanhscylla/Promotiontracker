import streamlit as pd
import streamlit as st
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
from datetime import datetime
import re
import io

# Cấu hình giao diện Streamlit (Slate Blue Professional Theme)
st.set_page_config(
    page_title="MWG SKU Promotion Tracking System",
    page_icon="📊",
    layout="wide"
)

# Thêm CSS custom để giao diện mang phong cách Dashboard chuyên nghiệp
st.markdown("""
    <style>
    .main-header { font-size:24px; font-weight:bold; color:#1F385C; margin-bottom:20px; }
    .stButton>button { background-color: #1F385C; color: white; border-radius: 4px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">📊 HỆ THỐNG THEO DÕI & CẬP NHẬT GIÁ SKU TỰ ĐỘNG</div>', unsafe_allow_html=True)

def clean_price(price_str):
    if not price_str:
        return 0
    price_num = re.sub(r'\D', '', price_str)
    return int(price_num) if price_num else 0

def tracking_promotion_detail(data_frame):
    # Cấu hình Chrome Chạy ẩn (Bắt buộc cho môi trường Web Server)
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    final_results = []

    # Tạo một thanh trạng thái động trên giao diện Streamlit
    status_text = st.empty()
    progress_bar = st.progress(0)
    total_rows = len(data_frame)

    try:
        for index, row in data_frame.iterrows():
            url = row['URL']
            model = row['Model']
            
            # Cập nhật trạng thái xử lý từng SKU lên màn hình web
            status_text.markdown(f"⏳ **Đang quét ({index+1}/{total_rows}):** {model}...")
            progress_bar.progress((index + 1) / total_rows)

            try:
                driver.get(url)
                time.sleep(3)

                goc = 0
                km = 0
                km_co_so_raw = 0
                thoi_han = "Không có"
                loai_hinh = "Thường"

                # 1. KIỂM TRA LOẠI HÌNH
                try:
                    bs_title = driver.find_element(By.CLASS_NAME, "bs_title")
                    if "Online" in bs_title.text:
                        loai_hinh = "Online"
                except:
                    pass

                # 2. LẤY GIÁ
                try:
                    service_pack = driver.find_element(By.CLASS_NAME, "box_servicepack")
                    first_pack = service_pack.find_element(By.CSS_SELECTOR, "[data-id='1'] span")
                    b_tags = first_pack.find_elements(By.TAG_NAME, "b")
                    b_price_text = next((b.text for b in b_tags if re.search(r'\d', b.text)), "")
                    em_tags = first_pack.find_elements(By.TAG_NAME, "em")
                    
                    if not em_tags or "hide" in em_tags[0].get_attribute("class") or not re.search(r'\d', em_tags[0].text):
                        km = clean_price(b_price_text)
                        goc = km
                    else:
                        goc = clean_price(em_tags[0].text)
                        km = clean_price(b_price_text)
                except:
                    try:
                        bs_price = driver.find_element(By.CLASS_NAME, "bs_price")
                        goc_attr = bs_price.get_attribute("data-priceOrg")
                        goc = int(float(goc_attr)) if goc_attr else 0
                        km_text = driver.find_element(By.CLASS_NAME, "box-price-present").text
                        km = clean_price(km_text)
                        loai_hinh = "Online"
                    except:
                        try:
                            box_price = driver.find_element(By.CLASS_NAME, "box-price")
                            km_text = box_price.find_element(By.CLASS_NAME, "box-price-present").text
                            goc_attr = box_price.get_attribute("data-priceOrg")
                            km = clean_price(km_text)
                            goc = int(float(goc_attr)) if goc_attr else km
                        except:
                            pass

                # 2.5 CÀO DỮ LIỆU KM CƠ SỞ
                try:
                    choose_promo = driver.find_element(By.CLASS_NAME, "choosepromo")
                    discount_attr = choose_promo.get_attribute("data-discount")
                    km_co_so_raw = int(float(discount_attr)) if discount_attr else 0
                except:
                    km_co_so_raw = 0

                # 3. LẤY THỜI HẠN KM
                try:
                    thoi_han = driver.find_element(By.CLASS_NAME, "pr-txt").text.replace("Giá và khuyến mãi dự kiến áp dụng đến", "").strip()
                except:
                    thoi_han = "Theo chương trình"

                # 4. TÍNH TOÁN
                gia_tri_km = goc - km if goc > km else 0
                phan_tram_km = round((gia_tri_km / goc) * 100, 1) if goc > 0 else 0
                km_co_so_final = km_co_so_raw if km_co_so_raw > 0 else gia_tri_km

                final_results.append({
                    "Ngày quét": datetime.now().strftime("%d/%m/%Y"),
                    "Model": model,
                    "Giá gốc": goc,
                    "Giá khuyến mãi": km,
                    "Giá trị KM cơ sở": km_co_so_final,
                    "Giá trị KM": gia_tri_km,
                    "% KM": phan_tram_km,
                    "Thời hạn KM": thoi_han,
                    "Hình thức": loai_hinh,
                    "Link": url,
                })
            except Exception as e:
                st.warning(f"Lỗi khi cào sản phẩm {model}: {e}")
    finally:
        driver.quit()
        status_text.success("🎉 Đã hoàn thành quét toàn bộ danh sách SKU!")

    return pd.DataFrame(final_results)

# ── THIẾT KẾ GIAO DIỆN CHỨC NĂNG ────────────────────────────────
st.sidebar.header("📁 Cấu Hình Dữ Liệu")
uploaded_file = st.sidebar.file_uploader("Tải lên file SKU Tracking (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    df_input = pd.read_excel(uploaded_file)
    st.subheader("📋 Danh sách SKU đầu vào cần theo dõi")
    st.dataframe(df_input, use_container_width=True)
    
    if st.button("🚀 Bắt đầu quét dữ liệu"):
        with st.spinner("Hệ thống đang chạy ngầm trình duyệt để cập nhật giá..."):
            df_result = tracking_promotion_detail(df_input)
            
            st.subheader("📊 Kết quả cập nhật thời gian thực")
            
            # Định dạng hiển thị tiền tệ trực quan trên bảng Streamlit
            st.dataframe(
                df_result.style.format({
                    "Giá gốc": "{:,.0f} đ",
                    "Giá khuyến mãi": "{:,.0f} đ",
                    "Giá trị KM cơ sở": "{:,.0f} đ",
                    "Giá trị KM": "{:,.0f} đ",
                    "% KM": "{:.1f}%"
                }), 
                use_container_width=True
            )
            
            # Xử lý xuất file xuất ra dạng Bytes để Streamlit Download trực tiếp
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_result.to_excel(writer, index=False)
            processed_data = output.getvalue()
            
            # Tạo nút tải file động phân loại Weekend / Normal
            now = datetime.now()
            suffix = "_weekend" if now.weekday() in [4, 5, 6] else "_normal"
            file_name = f"Bao_cao_Promotion_{now.strftime('%d_%m_%Y')}{suffix}.xlsx"
            
            st.download_button(
                label="📥 Tải xuống Báo Cáo Excel",
                data=processed_data,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("💡 Vui lòng tải file `SKU Tracking.xlsx` từ thanh Sidebar bên trái để bắt đầu.")