import rasterio
import numpy as np
import matplotlib.pyplot as plt
from skimage.segmentation import slic, mark_boundaries
import pandas as pd
import os
import sys
from glob import glob

# --- 1. TỰ ĐỘNG CHUYỂN VỀ THƯ MỤC CHỨA CODE (FIX LỖI ĐƯỜNG DẪN) ---
# Dù bạn chạy code từ đâu, Python cũng sẽ tự nhảy vào đúng thư mục chứa file main.py này
os.chdir(os.path.dirname(os.path.abspath(__file__)))
print(f"📂 Đang làm việc tại: {os.getcwd()}")

# --- 2. TỰ ĐỘNG TÌM THƯ MỤC DỮ LIỆU ---
found_folders = glob("*.SAFE")
if len(found_folders) == 0:
    print("❌ LỖI: Không tìm thấy thư mục .SAFE nào cả!")
    print("👉 Hãy kiểm tra: Thư mục .SAFE có đang nằm CÙNG CHỖ với file main.py không?")
    exit()

SAFE_NAME = found_folders[0]
print(f"✅ Đã tìm thấy dữ liệu: {SAFE_NAME}")

# --- 3. HÀM TÌM FILE ẢNH (TÌM KIẾM MẠNH MẼ HƠN) ---
def get_band(color_code):
    # Thay đổi chiến thuật: Tìm file chứa mã màu (VD: *B04*) bất kể có gạch dưới hay không
    search_pattern = os.path.join(SAFE_NAME, "**", f"*{color_code}*.jp2")
    files = glob(search_pattern, recursive=True)
    
    if not files:
        print(f"⚠️ Không tìm thấy file chứa mã '{color_code}'")
        return None

    # Lọc thông minh:
    # 1. Bỏ qua các file ảnh xem trước (thường có chữ PVI hoặc TCI)
    # 2. Ưu tiên file nằm trong thư mục IMG_DATA nếu có nhiều bản
    target_file = files[0]
    for f in files:
        if "IMG_DATA" in f and "TCI" not in f and "PVI" not in f:
            target_file = f
            break # Lấy được file chuẩn rồi thì dừng
            
    print(f"   -> Đọc file: {os.path.basename(target_file)}")
    
    with rasterio.open(target_file) as src:
        # Thu nhỏ 10 lần để chạy nhanh (Sentinel gốc 10000x10000 rất nặng)
        return src.read(1, out_shape=(src.height//10, src.width//10)).astype(float)

# --- CHƯƠNG TRÌNH CHÍNH ---
print("⏳ Đang đọc dữ liệu vệ tinh...")
try:
    red   = get_band("B04") 
    green = get_band("B03") 
    nir   = get_band("B08") 
except Exception as e:
    print(f"❌ Lỗi hệ thống: {e}"); exit()

if red is not None:
    # --- 4. TÍNH TOÁN & SUPER-PIXEL ---
    print("⚙️ Đang tính toán NDVI, NDWI và Phân lô...")
    
    NDVI = (nir - red) / (nir + red + 1e-6)
    NDWI = (green - nir) / (green + nir + 1e-6)
    
    Img = np.dstack((nir, red, green)) 
    Img = np.clip(Img / 3000.0, 0, 1)

    segments = slic(Img, n_segments=300, compactness=20, start_label=0)
    num_lots = len(np.unique(segments))
    
    print(f"✅ Đã chia bản đồ thành {num_lots} lô đất.")

    # --- 5. XUẤT EXCEL ---
    excel_filename = "Ket_Qua_Phan_Tich.xlsx"
    print(f"💾 Đang xuất dữ liệu ra file Excel: {excel_filename}")
    
    data_list = []

    for i in range(num_lots):
        mask = (segments == i)
        area = np.sum(mask)
        if area == 0 or np.mean(red[mask]) == 0: continue

        avg_ndvi = np.mean(NDVI[mask])
        avg_ndwi = np.mean(NDWI[mask])
        
        # Logic tính toán
        Wi = 1 if avg_ndwi > 0.0 else 0
        
        land_type = "Nước"
        if Wi == 0:
            if avg_ndvi > 0.4: land_type = "Rừng"
            elif avg_ndvi < 0.1: land_type = "Đất trống"
            else: land_type = "Cây bụi"

        unit_cost = 200 if land_type == "Rừng" else (50 if land_type == "Đất trống" else 100)
        Ci = area * unit_cost
        efficiency = 0 if Wi == 1 else (1.0 - max(0, avg_ndvi))
        Ei = area * 10 * efficiency

        data_list.append({
            "ID_Lô": i,
            "Loại_Đất": land_type,
            "Wi_Là_Nước": Wi,
            "Ei_Điện_Năng": round(Ei, 2),
            "Ci_Chi_Phí": int(Ci),
            "Diện_Tích": int(area),
            "NDVI": round(avg_ndvi, 3),
            "NDWI": round(avg_ndwi, 3)
        })

    # Lưu ra Excel
    df = pd.DataFrame(data_list)
    df.to_excel(excel_filename, index=False)

    print("🎉 XONG! Hãy mở file Excel lên xem nhé.")

    # Hiển thị
    plt.figure(figsize=(10, 8))
    plt.imshow(Img)
    plt.imshow(mark_boundaries(np.zeros_like(Img), segments, color=(1,1,0)), alpha=0.5)
    plt.title(f"Phân lô tự động ({os.path.basename(SAFE_NAME)})")
    plt.axis('off')
    plt.show()
else:
    print("❌ Không đọc được dữ liệu ảnh nào cả.")