import streamlit as st
import pandas as pd
import numpy as np
import random
import re
import traceback
import io

st.set_page_config(page_title="Hệ thống Nhật ký V15.0", layout="wide")
st.title("🚛 HỆ THỐNG TÍNH TOÁN NHẬT KÝ & PHIẾU CẤP (V15.0)")

# --- HÀM TÌM CỘT BẮT BUỘC ---
def req_c(df, keys, table_name):
    for c in df.columns:
        if any(k.lower() in str(c).lower() for k in keys): return c
    raise ValueError(f"❌ Không tìm thấy cột chứa từ khóa {keys} trong bảng {table_name}")

# --- HÀM TÁCH SỐ TỪ CHUỖI (VD: "80 - 100" -> 80, 100) ---
def parse_range(val, def_min=0.0, def_max=0.0):
    if pd.isna(val): return def_min, def_max
    nums = re.findall(r"[\d\.]+", str(val).replace(',', '.'))
    if len(nums) >= 2: return float(nums[0]), float(nums[1])
    elif len(nums) == 1: return float(nums[0]), float(nums[0])
    return def_min, def_max

# --- HÀM FORMAT GIAO DIỆN ---
def format_vietnamese_style(df):
    if df.empty: return df
    df_display = df.copy()
    df_display.index = np.arange(1, len(df_display) + 1)
    for col in df_display.columns:
        if pd.api.types.is_datetime64_any_dtype(df_display[col]):
            df_display[col] = df_display[col].dt.strftime('%d/%m/%Y')
    format_dict = {}
    for col in df_display.select_dtypes(include=['float64', 'int64']).columns:
        col_l = str(col).lower()
        if any(k in col_l for k in ['lượng', 'lít', 'cân', 'tích', 'thụ', 'km', 'giờ', 'thiếu']):
            format_dict[col] = "{:,.2f}"
        else:
            format_dict[col] = "{:,.0f}"
    return df_display.style.format(format_dict, decimal=',', thousands='.', na_rep='')

# --- GIAO DIỆN CHÍNH ---
with st.sidebar:
    st.header("⚙️ THAM SỐ TÍNH TOÁN")
    avg_velocity = st.slider("Vận tốc di chuyển TB (km/h)", 10.0, 60.0, 25.0, 1.0)
    tank_limit_pct = st.slider("Giới hạn đổ đầy bình (%)", 80.0, 100.0, 95.0, 1.0)
    st.info("💡 Thời gian làm việc = 8h - (Quãng đường / Vận tốc).")

uploaded_file = st.file_uploader("Nạp file Excel dữ liệu (Gồm 6 sheet)", type=["xlsx"])

if uploaded_file:
    try:
        # Nạp dữ liệu
        xls = pd.ExcelFile(uploaded_file)
        sheets = ['hoadon', 'tien_do', 'day_chuyen', 'dinh_muc', 'phieu_can', 'ngaynghi']
        data = {s: pd.read_excel(xls, sheet_name=s) for s in sheets}
        for s in data: data[s].columns = [str(c).strip() for c in data[s].columns]

        # 1. Hóa đơn
        hd = data['hoadon']
        c_hd_ngay = req_c(hd, ['ngày', 'thời gian'], 'Hóa đơn')
        c_hd_sp = req_c(hd, ['tên hàng', 'sản phẩm', 'hàng hóa'], 'Hóa đơn')
        c_hd_sl = req_c(hd, ['số lượng', 'lít'], 'Hóa đơn')
        c_hd_so = req_c(hd, ['hóa đơn', 'số hđ'], 'Hóa đơn')
        hd['dt_norm'] = pd.to_datetime(hd[c_hd_ngay], dayfirst=True, errors='coerce')

        # Tìm tên chính xác của Xăng và Dầu từ bảng Hóa đơn để đồng bộ
        ten_xang_chuan = "Xăng (Chưa xác định)"
        xang_samples = hd[hd[c_hd_sp].astype(str).str.lower().str.contains('xăng')][c_hd_sp].unique()
        if len(xang_samples) > 0: ten_xang_chuan = str(xang_samples[0])

        # 2. Tiến độ
        td = data['tien_do']
        c_td_tuyen = req_c(td, ['tuyến'], 'Tiến độ')
        c_td_dc = req_c(td, ['dây chuyền'], 'Tiến độ')
        td['start'] = pd.to_datetime(td[req_c(td, ['bắt đầu'], 'Tiến độ')], errors='coerce')
        td['end'] = pd.to_datetime(td[req_c(td, ['kết thúc'], 'Tiến độ')], errors='coerce')

        # 3. Dây chuyền
        dc = data['day_chuyen']
        c_dc_ma = req_c(dc, ['mã dây chuyền', 'mã hiệu', 'mã'], 'Dây chuyền')
        c_dc_tb = req_c(dc, ['danh mục', 'thiết bị'], 'Dây chuyền')

        # 4. Định mức
        dm = data['dinh_muc']
        c_dm_bs = req_c(dm, ['biển số', 'mã tb'], 'Định mức')
        c_dm_loai = req_c(dm, ['loại thiết bị'], 'Định mức')
        c_dm_nl = req_c(dm, ['loại nhiên liệu', 'nhiên liệu'], 'Định mức')
        c_dm_dungtich = req_c(dm, ['dung tích'], 'Định mức')
        c_dm_km = req_c(dm, ['100km'], 'Định mức')
        c_dm_range_km = req_c(dm, ['khoảng quãng đường', 'khoảng km'], 'Định mức')
        c_dm_h = req_c(dm, ['lít/giờ', 'l/h', 'giờ'], 'Định mức')
        c_dm_range_h = req_c(dm, ['khoảng thời gian'], 'Định mức')
        c_dm_ton = req_c(dm, ['tồn đầu kỳ'], 'Định mức')

        # 5. Phiếu cân
        pc = data['phieu_can']
        c_pc_ngay = req_c(pc, ['ngày'], 'Phiếu cân')
        c_pc_bs = req_c(pc, ['biển số'], 'Phiếu cân')
        pc['dt'] = pd.to_datetime(pc[c_pc_ngay], errors='coerce')
        pc[c_pc_bs] = pc[c_pc_bs].astype(str).str.strip().str.upper()

        # 6. Ngày nghỉ
        nn = data['ngaynghi']
        nn['dt'] = pd.to_datetime(nn[req_c(nn, ['ngày'], 'Ngày nghỉ')], errors='coerce')
        ngay_nghi_list = nn['dt'].dropna().dt.date.tolist()

        # Tiền xử lý chữ (Ép kiểu IN HOA và in thường)
        dm[c_dm_bs] = dm[c_dm_bs].astype(str).str.strip().str.upper()
        dm[c_dm_loai] = dm[c_dm_loai].astype(str).str.strip().str.lower()
        dc[c_dc_ma] = dc[c_dc_ma].astype(str).str.strip().str.upper()

        st.success("✅ Dữ liệu đã chuẩn hóa 100%. Đã nhận đủ 6 Sheet.")
        if st.button("🚀 TẠO NHẬT KÝ & PHIẾU CẤP", type="primary", use_container_width=True):
            tank_balances = {}
            kho_xang_balance = 0.0  # Tồn kho lũy kế của Can/Phuy Xăng
            
            res_nk = []
            res_phieu_cap = []
            res_thieu_hut = []
            unused_do_invoices = []
            
            # Tìm dải ngày thi công lớn nhất
            min_date = td['start'].min()
            max_date = td['end'].max()
            if pd.isnull(min_date) or pd.isnull(max_date):
                st.error("❌ Không tìm thấy dữ liệu ngày hợp lệ trong bảng Tiến độ.")
                st.stop()
                
            all_dates = pd.date_range(start=min_date, end=max_date)

            for current_date in all_dates:
                if current_date.date() in ngay_nghi_list: continue
                d_str = current_date.strftime('%d/%m/%Y')
                
                active_routes = td[(td['start'] <= current_date) & (td['end'] >= current_date)]
                if active_routes.empty: continue
                
                pc_today = pc[pc['dt'].dt.date == current_date.date()][c_pc_bs].unique()
                used_today = [] 
                daily_eqs = {}
                
                # --- A. GOM THIẾT BỊ VÀ TÍNH TIÊU THỤ (Chạy đủ 8 tiếng) ---
                for _, route in active_routes.iterrows():
                    dc_list = [x.strip().upper() for x in str(route[c_td_dc]).split(',') if x.strip()]
                    for dc_code in dc_list:
                        dc_row = dc[dc[c_dc_ma] == dc_code]
                        if dc_row.empty: continue
                        
                        tb_list = [x.strip().lower() for x in str(dc_row.iloc[0][c_dc_tb]).split(',') if x.strip()]
                        for tb_type in tb_list:
                            pool = dm[(dm[c_dm_loai] == tb_type) & (~dm[c_dm_bs].isin(used_today))]
                            if pool.empty: continue
                            
                            match_pc = pool[pool[c_dm_bs].isin(pc_today)]
                            selected_eq = match_pc.iloc[0] if not match_pc.empty else pool.iloc[0]
                            
                            plate = selected_eq[c_dm_bs]
                            used_today.append(plate)
                            
                            if plate not in tank_balances:
                                val_ton = selected_eq[c_dm_ton]
                                if pd.isna(val_ton): tank_balances[plate] = selected_eq[c_dm_dungtich] * 0.05
                                else: tank_balances[plate] = float(val_ton)
                            
                            min_km, max_km = parse_range(selected_eq[c_dm_range_km])
                            min_h, max_h = parse_range(selected_eq[c_dm_range_h])
                            
                            if max_km > 0:
                                km = round(random.uniform(min_km, max_km), 2)
                                t_move = km / avg_velocity
                                t_work = max(0.0, 8.0 - t_move)
                            else:
                                km = 0.0
                                t_work = round(random.uniform(min_h, max_h), 2)
                                
                            dm_km = float(selected_eq[c_dm_km]) if pd.notnull(selected_eq[c_dm_km]) else 0.0
                            dm_h = float(selected_eq[c_dm_h]) if pd.notnull(selected_eq[c_dm_h]) else 0.0
                            
                            daily_eqs[plate] = {
                                'tuyen': route[c_td_tuyen], 'dc': dc_code, 'loai': selected_eq[c_dm_loai].title(),
                                'loai_nl': str(selected_eq[c_dm_nl]), 'dung_tich': selected_eq[c_dm_dungtich],
                                't_dau': tank_balances[plate], 'km': km, 't_work': t_work,
                                'tt_du_kien': (km * dm_km / 100.0) + (t_work * dm_h)
                            }

                # --- B. XỬ LÝ HÓA ĐƠN & RÓT NHIÊN LIỆU (Cuối ngày) ---
                hd_today = hd[hd['dt_norm'].dt.date == current_date.date()]
                cap_ngay = {p: 0.0 for p in daily_eqs}
                
                hd_xang = hd_today[hd_today[c_hd_sp].astype(str).str.lower().str.contains('xăng')]
                hd_do = hd_today[~hd_today[c_hd_sp].astype(str).str.lower().str.contains('xăng')]
                
                # 1. Nhập Xăng vào KHO ẢO
                ton_kho_xang_dau_ngay = kho_xang_balance
                nhap_kho_xang = 0.0
                for _, row in hd_xang.iterrows():
                    sl_xang = row[c_hd_sl]
                    kho_xang_balance += sl_xang
                    nhap_kho_xang += sl_xang
                    res_phieu_cap.append({
                        'Ngày': d_str, 'Thiết bị nhận': 'Kho Can/Phuy Xăng', 
                        'Loại NL': row[c_hd_sp], 'Số lượng (L)': sl_xang,
                        'Căn cứ HĐ': row[c_hd_so], 'Ghi chú': 'Nhập Hóa đơn Xăng vào Kho'
                    })

                # 2. Rút Xăng từ Kho Ảo đổ ĐẦY BÌNH cho Máy Nhỏ (Giới hạn theo tồn kho)
                xuat_kho_xang = 0.0
                xe_xang = [p for p in daily_eqs if 'xăng' in str(daily_eqs[p]['loai_nl']).lower()]
                for plate in xe_xang:
                    info = daily_eqs[plate]
                    t_dau = info['t_dau']
                    tt = info['tt_du_kien']
                    
                    # Nhu cầu đổ xăng = (Dung tích * %) - Tồn hiện tại sau khi chạy
                    ton_hien_tai_sau_chay = t_dau - tt
                    if ton_hien_tai_sau_chay <= 0:  # Máy bị cạn hoặc âm
                        nhu_cau = (info['dung_tich'] * tank_limit_pct / 100.0) - ton_hien_tai_sau_chay
                        luong_cap = min(nhu_cau, kho_xang_balance)  # Rót tối đa bằng số xăng đang có
                        
                        if luong_cap > 0:
                            cap_ngay[plate] = luong_cap
                            kho_xang_balance -= luong_cap
                            xuat_kho_xang += luong_cap
                            res_phieu_cap.append({
                                'Ngày': d_str, 'Thiết bị nhận': plate, 
                                'Loại NL': info['loai_nl'], 'Số lượng (L)': luong_cap,
                                'Căn cứ HĐ': 'Xuất nội bộ', 'Ghi chú': 'Châm bình từ Can/Phuy'
                            })

                # Chốt Nhật ký cho thiết bị Ảo "Kho Can Xăng"
                res_nk.append({
                    'Ngày': d_str, 'Tuyến': 'Quản lý Nội bộ', 'Biển số / Mã TB': 'KHO-CAN-XANG',
                    'Loại TB': 'Kho chứa Xăng', 'Nhiên liệu': ten_xang_chuan,
                    'Tồn đầu (L)': ton_kho_xang_dau_ngay, 'Số cấp (L)': nhap_kho_xang,
                    'Km di chuyển': 0.0, 'Giờ tại chỗ': 0.0, 'Tiêu thụ (L)': xuat_kho_xang,
                    'Thiếu hụt (L)': 0.0, 'Tồn cuối (L)': kho_xang_balance
                })

                # 3. Thuật toán Bin-Packing 2 Vòng cho Hóa đơn DO
                # Sắp xếp hóa đơn từ TO đến NHỎ
                hd_do_list = hd_do.sort_values(by=c_hd_sl, ascending=False).to_dict('records')
                used_hd = []
                xe_do = [p for p in daily_eqs if 'dầu' in str(daily_eqs[p]['loai_nl']).lower()]
                
                # Tính khoảng trống hiện tại (sau khi đã trừ tiêu thụ trong ngày)
                for plate in xe_do:
                    daily_eqs[plate]['khoang_trong'] = (daily_eqs[plate]['dung_tich'] * tank_limit_pct / 100.0) - (daily_eqs[plate]['t_dau'] - daily_eqs[plate]['tt_du_kien'])
                    daily_eqs[plate]['da_nhan_hd'] = False  # Cờ đánh dấu Vòng 1
                
                # VÒNG 1: Ưu tiên công bằng (Mỗi xe tối đa 1 hóa đơn)
                # Sắp xếp xe theo khoảng trống từ TO đến NHỎ
                xe_do = sorted(xe_do, key=lambda p: daily_eqs[p]['khoang_trong'], reverse=True)
                for inv in hd_do_list:
                    if inv[c_hd_so] in used_hd: continue
                    for plate in xe_do:
                        if not daily_eqs[plate]['da_nhan_hd'] and inv[c_hd_sl] <= daily_eqs[plate]['khoang_trong']:
                            cap_ngay[plate] += inv[c_hd_sl]
                            daily_eqs[plate]['khoang_trong'] -= inv[c_hd_sl]
                            daily_eqs[plate]['da_nhan_hd'] = True
                            used_hd.append(inv[c_hd_so])
                            res_phieu_cap.append({
                                'Ngày': d_str, 'Thiết bị nhận': plate, 
                                'Loại NL': inv[c_hd_sp], 'Số lượng (L)': inv[c_hd_sl],
                                'Căn cứ HĐ': inv[c_hd_so], 'Ghi chú': 'Đổ trực tiếp (Vòng 1 - Công bằng)'
                            })
                            break # Đã ghép xong tờ HĐ này
                
                # VÒNG 2: Tích trữ (Ai còn trống thì nhét nốt hóa đơn thừa)
                # Cập nhật lại danh sách xe theo khoảng trống mới nhất (TO đến NHỎ)
                xe_do = sorted(xe_do, key=lambda p: daily_eqs[p]['khoang_trong'], reverse=True)
                for inv in hd_do_list:
                    if inv[c_hd_so] in used_hd: continue
                    for plate in xe_do:
                        if inv[c_hd_sl] <= daily_eqs[plate]['khoang_trong']:
                            cap_ngay[plate] += inv[c_hd_sl]
                            daily_eqs[plate]['khoang_trong'] -= inv[c_hd_sl]
                            used_hd.append(inv[c_hd_so])
                            res_phieu_cap.append({
                                'Ngày': d_str, 'Thiết bị nhận': plate, 
                                'Loại NL': inv[c_hd_sp], 'Số lượng (L)': inv[c_hd_sl],
                                'Căn cứ HĐ': inv[c_hd_so], 'Ghi chú': 'Đổ trực tiếp (Vòng 2 - Tích trữ)'
                            })
                            break # Đã ghép xong tờ HĐ này

                # Gom hóa đơn DO dư thừa để cảnh báo
                for inv in hd_do_list:
                    if inv[c_hd_so] not in used_hd:
                        unused_do_invoices.append({
                            'Ngày': d_str, 'Số HĐ': inv[c_hd_so], 
                            'Sản phẩm': inv[c_hd_sp], 'Số lượng (Lít)': inv[c_hd_sl]
                        })

                # --- C. CHỐT SỐ LIỆU (Giữ nguyên 8 tiếng, Ghi nhận Nợ) ---
                for plate, info in daily_eqs.items():
                    tong_dau_hien_co = info['t_dau'] + cap_ngay[plate]
                    tt_thuc = info['tt_du_kien']
                    thieu_hut = 0.0
                    ton_cuoi_tam = tong_dau_hien_co - tt_thuc
                    
                    # Nếu xe/máy bị âm dầu (chạy không đủ)
                    if ton_cuoi_tam < 0:
                        thieu_hut = abs(ton_cuoi_tam)
                        ton_cuoi_tam = 0.0
                        res_thieu_hut.append({
                            'Ngày': d_str, 'Biển số / Mã TB': plate,
                            'Nhiên liệu cần': info['loai_nl'], 'Nhiên liệu đang có': tong_dau_hien_co,
                            'Số Lít Thiếu Hụt': thieu_hut, 'Ghi chú': 'Không đủ nhiên liệu, chạy nợ'
                        })

                    tank_balances[plate] = ton_cuoi_tam
                    
                    res_nk.append({
                        'Ngày': d_str, 'Tuyến': info['tuyen'], 'Biển số / Mã TB': plate,
                        'Loại TB': info['loai'], 'Nhiên liệu': info['loai_nl'],
                        'Tồn đầu (L)': info['t_dau'], 'Số cấp (L)': cap_ngay[plate],
                        'Km di chuyển': info['km'], 'Giờ tại chỗ': info['t_work'],
                        'Tiêu thụ (L)': tt_thuc, 'Thiếu hụt (L)': thieu_hut,
                        'Tồn cuối (L)': tank_balances[plate]
                    })

            # --- HIỂN THỊ KẾT QUẢ TỔNG HỢP ---
            st.divider()
            
            # CẢNH BÁO 1: HÓA ĐƠN DO DƯ THỪA
            if unused_do_invoices:
                st.error("⚠️ CẢNH BÁO 1: HÓA ĐƠN DẦU DO DƯ THỪA TRONG NGÀY (Không xe nào chứa vừa)")
                st.dataframe(pd.DataFrame(unused_do_invoices), use_container_width=True)

            # CẢNH BÁO 2: THIẾT BỊ ĐÓI NHIÊN LIỆU
            if res_thieu_hut:
                st.warning("⚠️ CẢNH BÁO 2: DANH SÁCH THIẾT BỊ BỊ THIẾU HỤT NHIÊN LIỆU TRONG NGÀY")
                st.dataframe(pd.DataFrame(res_thieu_hut), use_container_width=True)

            if res_nk:
                t1, t2 = st.tabs(["📊 BẢNG 1: NHẬT KÝ THI CÔNG", "📑 BẢNG 2: PHIẾU CẤP"])
                df_nk = pd.DataFrame(res_nk)
                df_pc = pd.DataFrame(res_phieu_cap)
                
                with t1: st.dataframe(format_vietnamese_style(df_nk), height=600, use_container_width=True)
                with t2: st.dataframe(format_vietnamese_style(df_pc), height=600, use_container_width=True)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_nk.to_excel(writer, index=False, sheet_name='NhatKy')
                    df_pc.to_excel(writer, index=False, sheet_name='PhieuCap')
                    if res_thieu_hut: pd.DataFrame(res_thieu_hut).to_excel(writer, index=False, sheet_name='CB_ThieuHut')
                    if unused_do_invoices: pd.DataFrame(unused_do_invoices).to_excel(writer, index=False, sheet_name='CB_HoaDonDu')
                st.download_button("📥 Tải Xuống File Excel Tổng Hợp (V15.0)", output.getvalue(), "Ket_Qua_V15.xlsx")
            else:
                st.warning("⚠️ Không phát sinh dữ liệu nhật ký.")

    except Exception as e:
        st.error(f"❌ Lỗi hệ thống: {e}")
        st.code(traceback.format_exc())