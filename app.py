# -*- coding: utf-8 -*-
"""
프로젝트명: KCS 딸깍 기록지
제작자: D
슬로건: 당신의 손목은 소중하니깐..
버전: v5.5 (온도기록계 다채널/괄호, 압력 소수점 포맷 및 숨김 행/0값 처리 완벽 패치)
"""

import io
import os
import re
import openpyxl
from openpyxl.styles import Alignment
import streamlit as st

st.set_page_config(
    page_title="KCS 딸깍 기록지",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ==============================================================================
# [UI/UX 디자인 시스템] 반응형 모던 SaaS 인터페이스 & CSS 가변 레이아웃
# ==============================================================================
st.markdown(
    """
    <style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    
    html, body, [class*="css"], div:not([data-testid="stIconMaterial"]) {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif;
        box-sizing: border-box;
    }
    
    .stApp {
        background-color: #f7f9fa;
        background-image: 
            radial-gradient(at 0% 0%, rgba(199, 210, 254, 0.45) 0px, transparent 50%),
            radial-gradient(at 100% 0%, rgba(220, 252, 231, 0.55) 0px, transparent 50%),
            radial-gradient(at 50% 100%, rgba(254, 240, 138, 0.35) 0px, transparent 50%);
        background-attachment: fixed;
    }
    
    .block-container {
        max-width: 1200px !important;
        padding-top: 2.2rem !important;
        padding-bottom: 4rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
    }

    /* 메인 타이틀 */
    .hero-title-large {
        font-size: clamp(3.2rem, 7vw, 5.0rem);
        font-weight: 900;
        line-height: 1.1;
        color: #0F172A;
        letter-spacing: -0.04em;
        margin-bottom: 0.8rem;
    }
    
    .highlight-yellow {
        background: linear-gradient(180deg, transparent 62%, #BEF264 62%);
        display: inline;
        padding: 0 4px;
        border-radius: 4px;
    }
    
    /* 안내 서브 문구 */
    .hero-sub-notice {
        font-size: clamp(0.92rem, 1.6vw, 1.05rem);
        color: #475569;
        line-height: 1.65;
        font-weight: 500;
        margin-bottom: 1.6rem;
    }

    .badge-wrap {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        margin-bottom: 2.2rem;
    }
    .chip {
        display: inline-flex;
        align-items: center;
        padding: 0.38rem 0.85rem;
        border-radius: 9999px;
        font-size: 0.84rem;
        font-weight: 600;
        background: #ffffff;
        border: 1px solid #E2E8F0;
        box-shadow: 0 1px 2px rgba(0,0,0,0.02);
    }
    .chip-code {
        font-weight: 800;
        margin-right: 6px;
    }
    .chip-f { color: #EA580C; background: #FFF7ED; border-color: #FFEDD5; }
    .chip-m { color: #4F46E5; background: #EEF2FF; border-color: #E0E7FF; }
    .chip-t { color: #DB2777; background: #FDF2F8; border-color: #FCE7F3; }
    .chip-p { color: #059669; background: #ECFDF5; border-color: #D1FAE5; }

    /* 드롭존 커스텀 */
    div[data-testid="stFileUploader"] {
        width: 100% !important;
        margin-top: 0 !important;
        margin-bottom: 1.0rem !important;
    }
    div[data-testid="stFileUploader"] > label {
        display: none !important;
    }
    div[data-testid="stFileUploader"] section {
        position: relative !important;
        padding: 2.6rem 1.4rem 1.6rem 1.4rem !important;
        border: 2px dashed #94A3B8 !important;
        border-radius: 20px !important;
        background: rgba(255, 255, 255, 0.95) !important;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.03) !important;
        transition: all 0.2s ease !important;
    }
    div[data-testid="stFileUploader"] section::before {
        content: "교정 성적서를 끌어다 놓으세요";
        position: absolute;
        top: 14px;
        left: 0;
        right: 0;
        text-align: center;
        font-size: 1.05rem;
        font-weight: 800;
        color: #1E293B;
        letter-spacing: -0.02em;
        pointer-events: none;
    }
    div[data-testid="stFileUploader"] section:hover {
        border-color: #3B82F6 !important;
        background: #F0FDF4 !important;
    }
    div[data-testid="stFileUploader"] [data-testid="stIconMaterial"] {
        font-family: 'Material Symbols Rounded', 'Material Icons' !important;
    }
    div[data-testid="stFileUploader"] button {
        border-radius: 10px !important;
        font-size: 0.9rem !important;
    }
    
    /* 옵션 아이템 스타일 */
    .opt-icon-box {
        width: 38px;
        height: 38px;
        background: #ffffff;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.15rem;
        flex-shrink: 0;
    }
    .opt-main-lbl {
        font-size: 0.95rem;
        font-weight: 700;
        color: #0F172A;
        line-height: 1.2;
    }
    .opt-sub-lbl {
        font-size: 0.76rem;
        color: #64748B;
        margin-top: 2px;
    }
    
    .info-callout {
        background: #FEFCE8;
        border: 1px solid #FEF08A;
        border-radius: 12px;
        padding: 0.85rem 1rem;
        margin-top: 1.5rem;
        font-size: 0.8rem;
        color: #854D0E;
        line-height: 1.45;
        display: flex;
        gap: 8px;
        align-items: flex-start;
    }

    /* 하단 버튼 2개 통일 스타일링 */
    div.action-btn-del button {
        height: 48px !important;
        border: 1px solid #CBD5E1 !important;
        background: #FFFFFF !important;
        color: #475569 !important;
        font-size: 0.95rem !important;
        font-weight: 700 !important;
        border-radius: 12px !important;
        transition: all 0.2s ease !important;
    }
    div.action-btn-del button:hover {
        color: #EF4444 !important;
        background: #FEE2E2 !important;
        border-color: #FCA5A5 !important;
    }

    div.action-btn-run button {
        height: 48px !important;
        font-size: 0.95rem !important;
        font-weight: 700 !important;
        border-radius: 12px !important;
        transition: all 0.2s ease !important;
    }

    div.action-btn-run button:disabled {
        background-color: #FFFFFF !important;
        color: #94A3B8 !important;
        border: 1px solid #E2E8F0 !important;
        cursor: not-allowed !important;
        opacity: 0.75 !important;
        box-shadow: none !important;
    }

    /* 결과 메트릭 카드 모던화 */
    div[data-testid="stMetric"] {
        background: #ffffff !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 16px !important;
        padding: 0.85rem 1.0rem !important;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.03) !important;
        min-width: 0 !important;
        width: 100% !important;
    }
    div[data-testid="stMetricLabel"] {
        font-size: 0.85rem !important;
        font-weight: 600 !important;
        color: #64748B !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.45rem !important;
        font-weight: 800 !important;
        color: #0F172A !important;
        white-space: normal !important;
        word-break: break-all !important;
        line-height: 1.25 !important;
        margin-top: 0.2rem !important;
    }
    div[data-testid="stDownloadButton"] button {
        height: 48px !important;
        border-radius: 14px !important;
        font-size: 1.05rem !important;
        font-weight: 700 !important;
        padding: 0.75rem 1.5rem !important;
        box-shadow: 0 4px 14px rgba(37, 99, 235, 0.2) !important;
    }

    @media (max-width: 991px) {
        .block-container {
            padding-top: 1.5rem !important;
            padding-left: 1rem !important;
            padding-right: 1rem !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------- 상단 히어로 헤더 -----------------
st.markdown(
    """
    <div class="hero-title-large">
        <span class="highlight-yellow">KCS 딸깍 기록지</span>
    </div>
    <div class="hero-sub-notice">
        ⚠️ 변환 결과에 일부 오류가 있을 수 있으니 최종 확인을 권장드립니다.<br>
        업데이트 요청이나 오류 수정 문의는 <b>D</b>에게 전달해 주세요.
    </div>
    <div class="badge-wrap">
        <div class="chip chip-f"><span class="chip-code">FXX-</span> 힘 시험기</div>
        <div class="chip chip-m"><span class="chip-code">MXX-</span> 저울</div>
        <div class="chip chip-t"><span class="chip-code">TXX-</span> 온도 지시·기록계</div>
        <div class="chip chip-p"><span class="chip-code">PXX-</span> 압력계·차압계</div>
    </div>
    """,
    unsafe_allow_html=True,
)

TEMPLATE_FILE = "10칸 기록지.xlsx"

if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

# ----------------- 2열 가변 레이아웃 -----------------
col_left, col_right = st.columns([1.5, 1.0], gap="large")

with col_right:
    c_ico0, c_lbl0, c_sw0 = st.columns([0.8, 3.2, 1.2])
    with c_ico0:
        st.markdown('<div class="opt-icon-box" style="color: #6366F1;">⚡</div>', unsafe_allow_html=True)
    with c_lbl0:
        st.markdown('<div class="opt-main-lbl">자동 작업 모드</div><div class="opt-sub-lbl">ON: 즉시 변환 / OFF: 작업시작 클릭</div>', unsafe_allow_html=True)
    with c_sw0:
        auto_mode = st.toggle("auto_mode", value=True, label_visibility="collapsed")

    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

    c_ico1, c_lbl1, c_sw1 = st.columns([0.8, 3.2, 1.2])
    with c_ico1:
        st.markdown('<div class="opt-icon-box" style="color: #3B82F6;">#</div>', unsafe_allow_html=True)
    with c_lbl1:
        st.markdown('<div class="opt-main-lbl">넘버 삭제</div><div class="opt-sub-lbl">No 열 비우기</div>', unsafe_allow_html=True)
    with c_sw1:
        del_number = st.toggle("del_num", value=False, label_visibility="collapsed")

    st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

    c_ico2, c_lbl2, c_sw2 = st.columns([0.8, 3.2, 1.2])
    with c_ico2:
        st.markdown('<div class="opt-icon-box" style="color: #64748B;">📐</div>', unsafe_allow_html=True)
    with c_lbl2:
        st.markdown('<div class="opt-main-lbl">명목값 삭제</div><div class="opt-sub-lbl">범위 열 비우기</div>', unsafe_allow_html=True)
    with c_sw2:
        del_nominal = st.toggle("del_nom", value=False, label_visibility="collapsed")

    st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

    c_ico3, c_lbl3, c_sw3 = st.columns([0.8, 3.2, 1.2])
    with c_ico3:
        st.markdown('<div class="opt-icon-box" style="color: #10B981;">📊</div>', unsafe_allow_html=True)
    with c_lbl3:
        st.markdown('<div class="opt-main-lbl">데이터 삭제</div><div class="opt-sub-lbl">측정결과 열 비우기</div>', unsafe_allow_html=True)
    with c_sw3:
        del_data = st.toggle("del_dat", value=False, label_visibility="collapsed")

    st.markdown(
        """
        <div class="info-callout">
            <span>💡</span>
            <span>옵션을 켜면 해당 열은 빈칸으로 출력됩니다. 수기로 작성할 때 유용합니다.</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col_left:
    uploaded_files = st.file_uploader(
        "교정 성적서를 끌어다 놓으세요",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        key=f"file_uploader_{st.session_state['uploader_key']}",
        help=".xlsx, .xlsm 지원",
    )
    
    c_del, c_run = st.columns([1, 1], gap="medium")
    with c_del:
        st.markdown('<div class="action-btn-del">', unsafe_allow_html=True)
        if st.button("전체 삭제", use_container_width=True, help="업로드된 모든 파일을 즉시 비웁니다."):
            st.session_state["uploader_key"] += 1
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with c_run:
        st.markdown('<div class="action-btn-run">', unsafe_allow_html=True)
        if auto_mode:
            manual_run = st.button("⚡ 작업 시작 (자동 진행 중)", use_container_width=True, disabled=True)
        else:
            manual_run = st.button("⚡ 작업 시작", use_container_width=True, type="primary")
        st.markdown('</div>', unsafe_allow_html=True)

bottom_result_container = st.container()

# ==============================================================================
# [백엔드 비즈니스 로직 및 계산 함수]
# ==============================================================================

def get_decimal_places(val):
    if val is None or str(val).strip() == "" or str(val).strip() == "-":
        return 0
    try:
        s = str(val).strip()
        if "." in s:
            dec = s.split(".")[1]
            if len(dec) > 5 and ('9999' in dec or '0000' in dec):
                s = f"{round(float(s), 4)}"
                if "." in s:
                    return len(s.split(".")[1].rstrip('0'))
                return 0
            return len(dec)
    except Exception:
        pass
    return 0

def format_number(val, decimals):
    if val is None or val == "" or val == "-":
        return "-"
    try:
        f = float(val)
        if decimals == 0:
            return f"{int(round(f))}"
        else:
            fmt = "{:." + str(decimals) + "f}"
            return fmt.format(f)
    except Exception:
        return str(val).replace(",", "")

def format_meaningful_decimals(val):
    """표준값(범위)에서 무의미한 소수점(.00)을 제거하여 유의미한 값만 반환"""
    if val is None or str(val).strip() in ("", "-"):
        return "-"
    try:
        f_val = float(str(val).replace(",", "").strip())
        if f_val.is_integer():
            return str(int(f_val))
        else:
            s = f"{f_val:.6f}".rstrip('0').rstrip('.')
            return s
    except Exception:
        return str(val)

def get_formatted_value(cell):
    """지시값(측정결과)에서 Excel 셀 서식(number_format)을 분석하여 소수점을 그대로 유지"""
    if cell is None or cell.value is None or str(cell.value).strip() in ("", "-"):
        return None
    val = cell.value
    val_str = str(val).strip()
    
    if isinstance(val, (int, float)):
        fmt = str(cell.number_format)
        if '.' in fmt:
            dec_part = fmt.split('.')[1].split(';')[0]
            # 0, #, ? 문자의 개수를 세어 소수점 자릿수로 사용
            decimals = sum(1 for ch in dec_part if ch in '0#?')
            val_str = f"{val:.{decimals}f}"
        elif fmt != 'General' and float(val).is_integer():
            val_str = f"{int(val)}"
            
    return val_str

def format_sub_weight(val):
    if val is None or val == "" or val == "-":
        return "-"
    try:
        f = float(val)
        return f"{int(round(f))}"
    except Exception:
        return str(val).split(".")[0].replace(",", "")

def extract_set_count(text):
    if not text:
        return None
    m = re.search(r"(\d+)\s*조", str(text))
    return int(m.group(1)) if m else None

def clean_text(v):
    if v is None:
        return ""
    s = str(v).strip()
    if s.startswith("="):
        return ""
    return s

def format_capacity_and_resolution_robust(cap_raw, res_raw):
    c_str = str(cap_raw or "").strip()
    r_str = str(res_raw or "").strip()

    c_norm = re.sub(r'(\d+)\s+(\d{3})(?!\d)', r'\1\2', c_str)
    r_norm = re.sub(r'(\d+)\s+(\d{3})(?!\d)', r'\1\2', r_str)
    c_norm_check = re.sub(r'(\d+),(\d{3})(?!\d)', r'\1\2', c_norm)

    is_dual = False
    if (',' in r_norm or '/' in r_norm) and len(re.findall(r'[\d\.]+', r_norm)) >= 2:
        is_dual = True
    elif ',' in c_norm and len(re.findall(r'[\d\.]+', c_norm_check)) >= 2:
        is_dual = True

    if is_dual:
        cap_nums = re.findall(r'([\d\.]+)', c_norm)
        cap_units = re.findall(r'(kg|g|mg|t)', c_norm)
        unit_c = cap_units[-1] if cap_units else "g"
        if len(cap_nums) >= 2:
            c1 = f"{int(float(cap_nums[0])) if float(cap_nums[0]).is_integer() else cap_nums[0]}"
            c2 = f"{int(float(cap_nums[1])) if float(cap_nums[1]).is_integer() else cap_nums[1]}"
            max_str = f"{c1} / {c2} {unit_c}"
        else:
            max_str = c_norm

        res_nums = re.findall(r'([\d\.]+)', r_norm)
        res_units = re.findall(r'(kg|g|mg|t)', r_norm)
        unit_r = res_units[-1] if res_units else "mg"
        if len(res_nums) >= 2:
            r1 = f"{int(float(res_nums[0])) if float(res_nums[0]).is_integer() else res_nums[0]}"
            r2 = f"{int(float(res_nums[1])) if float(res_nums[1]).is_integer() else res_nums[1]}"
            div_str = f"{r1} / {r2} {unit_r}"
        else:
            div_str = r_norm
    else:
        cap_nums = re.findall(r'([\d\.]+)', c_norm_check)
        cap_units = re.findall(r'(kg|g|mg|t)', c_norm_check)
        unit_c = cap_units[-1] if cap_units else "kg"
        if cap_nums:
            f = float(cap_nums[0])
            max_str = f"{int(f) if f.is_integer() else f} {unit_c}"
        else:
            max_str = c_norm_check

        res_nums = re.findall(r'([\d\.]+)', r_norm)
        res_units = re.findall(r'(kg|g|mg|t)', r_norm)
        unit_r = res_units[-1] if res_units else "kg"
        if res_nums:
            f = float(res_nums[0])
            div_str = f"{int(f) if f.is_integer() else f} {unit_r}"
        else:
            div_str = r_norm

    return max_str, div_str, is_dual

def determine_scale_decimal_places(div_str, max_str, is_dual):
    if is_dual:
        c_units = re.findall(r'(kg|g|mg)', max_str)
        r_units = re.findall(r'(kg|g|mg)', div_str)
        u_c = c_units[-1] if c_units else "g"
        u_r = r_units[0] if r_units else "mg"
        r_nums = re.findall(r'([\d\.]+)', div_str)
        if r_nums:
            min_r = min([float(x) for x in r_nums])
            ratio = 1.0
            if u_r == "mg" and u_c == "g": ratio = 0.001
            elif u_r == "g" and u_c == "kg": ratio = 0.001
            elif u_r == "mg" and u_c == "kg": ratio = 0.000001
            actual = min_r * ratio
            s = f"{actual:.10f}".rstrip("0")
            if "." in s:
                return len(s.split(".")[1])
        return 0
    else:
        r_nums = re.findall(r'([\d\.]+)', div_str)
        if r_nums:
            s_val = r_nums[0]
            if "." in s_val:
                return len(s_val.split(".")[1])
        return 0

def safe_load_workbook(uploaded_file, data_only=True):
    uploaded_file.seek(0)
    try:
        return openpyxl.load_workbook(uploaded_file, data_only=data_only)
    except Exception:
        uploaded_file.seek(0)
        return openpyxl.load_workbook(uploaded_file, data_only=data_only, read_only=True)

# ==============================================================================
# [실행부: 엑셀 파싱 및 기록지 생성 로직]
# ==============================================================================

should_run = (auto_mode and uploaded_files) or (not auto_mode and uploaded_files and manual_run)

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"⚠️ 작업 폴더에 '{TEMPLATE_FILE}' 양식 파일이 준비되어 있어야 합니다.")
elif should_run:
    with st.spinner("⚡ 교정 성적서를 분석하고 10칸 기록지를 자동 생성하고 있습니다... 잠시만 기다려주세요."):
        try:
            all_instruments = []
            global_client_name = ""

            for file_idx, file in enumerate(uploaded_files, start=1):
                fname = file.name.upper()
                wb_src = safe_load_workbook(file, data_only=True)

                raw_device_name = ""
                client_name = ""
                cert_no = ""
                maker_name = ""
                model_name = ""
                serial_no = ""
                total_sets = None
                capacity_raw = ""
                scale_raw = ""

                # 기본정보 탭 확인
                info_sheet_names = [s for s in wb_src.sheetnames if "기본정보" in s and "영문" not in s]
                if info_sheet_names:
                    ws_info = wb_src[info_sheet_names[0]]
                    for r_idx, row in enumerate(ws_info.iter_rows(values_only=True)):
                        if r_idx > 15:
                            break
                        for val in row[:6]:
                            cell_v = str(val or "")
                            if "<Calibration_Certificate>" in cell_v:
                                m_c = re.search(r'<ClientInfo Name=[\'"]([^\'"]+)', cell_v)
                                if m_c: client_name = m_c.group(1)
                                m_dev = re.search(r'<InstrumentInfo[^>]*Description=[\'"]([^\'"]+)', cell_v)
                                if m_dev: raw_device_name = m_dev.group(1).strip()
                                m_m = re.search(r'<InstrumentInfo[^>]*Maker=[\'"]([^\'"]+)', cell_v)
                                if m_m: maker_name = m_m.group(1).strip()
                                m_mod = re.search(r'ManufacturerModel=[\'"]([^\'"]+)', cell_v)
                                if m_mod and "/" in m_mod.group(1):
                                    model_name = m_mod.group(1).split("/")[1].strip()
                                m_s = re.search(r'SerialNumber=[\'"]([^\'"]+)', cell_v)
                                if m_s: serial_no = m_s.group(1).strip()
                                m_cap = re.search(r'Capacity=[\'"]([^\'"]+)', cell_v)
                                if m_cap:
                                    capacity_raw = m_cap.group(1).strip()
                                    total_sets = extract_set_count(capacity_raw)
                                m_scale = re.search(r'Scale=[\'"]([^\'"]+)', cell_v)
                                if m_scale:
                                    scale_raw = m_scale.group(1).strip()

                    for r_idx, row in enumerate(ws_info.iter_rows(values_only=True)):
                        row_str = " ".join([str(c or "") for c in row[:35]])
                        if not client_name and "기 관 명" in row_str:
                            client_name = clean_text(row[6] if len(row) > 6 else "")
                        if not cert_no and "성적서번호" in row_str:
                            cert_no = clean_text(row[6] if len(row) > 6 else "")
                        if not maker_name and "제작회사" in row_str:
                            maker_name = clean_text(row[6] if len(row) > 6 else "")
                        if not raw_device_name and "기 기 명" in row_str:
                            raw_device_name = clean_text(row[6] if len(row) > 6 else "")
                        if not model_name and "형식·모델" in row_str:
                            val_m = clean_text(row[26] if len(row) > 26 and row[26] else (row[6] if len(row) > 6 else ""))
                            if val_m and val_m != "#VALUE!": model_name = val_m
                        if not capacity_raw and "측정범위" in row_str:
                            capacity_raw = clean_text(row[6] if len(row) > 6 else "")
                        if total_sets is None and ("측정범위" in row_str or "형식·모델" in row_str):
                            total_sets = extract_set_count(row[6] if len(row) > 6 else "")

                # 성적서 번호 폴백 탐색
                if not cert_no:
                    for sname in wb_src.sheetnames:
                        ws_temp = wb_src[sname]
                        for r in range(1, 15):
                            for c in range(1, min(25, ws_temp.max_column + 1)):
                                v_c = str(ws_temp.cell(r, c).value or "")
                                m_cert = re.search(r'([A-Za-z]\d{2}-\d+-\d+)', v_c)
                                if m_cert:
                                    cert_no = m_cert.group(1)
                                    break
                            if cert_no: break
                        if cert_no: break

                if client_name and not global_client_name:
                    global_client_name = client_name

                file_instruments = []

                # ========================================================
                # [기기 판별 A] 자동 호퍼 저울
                # ========================================================
                if "호퍼" in raw_device_name:
                    ws_rec = wb_src["기록지"]
                    all_rec_rows = list(ws_rec.iter_rows(values_only=True))

                    if not total_sets:
                        max_num = 0
                        for r_idx, r_vals in enumerate(all_rec_rows):
                            v60 = str(r_vals[59] if len(r_vals) > 59 else "").strip()
                            if v60.startswith("(") and v60.endswith(")") and v60[1:-1].isdigit():
                                if r_idx + 1 < len(all_rec_rows) and len(all_rec_rows[r_idx+1]) > 6 and all_rec_rows[r_idx+1][6] is not None:
                                    max_num = max(max_num, int(v60[1:-1]))
                        total_sets = max_num if max_num > 0 else 3

                    clean_dev = raw_device_name.replace("자동호퍼저울", "자동 호퍼 저울").strip()
                    dev_disp = f"{clean_dev}\n<{total_sets}조>"

                    for r_idx, r_vals in enumerate(all_rec_rows):
                        val_60 = str(r_vals[59] if len(r_vals) > 59 else "").strip()
                        if val_60.startswith("(") and val_60.endswith(")") and val_60[1:-1].isdigit():
                            num = int(val_60[1:-1])
                            if num > total_sets: continue
                            r_info_idx = r_idx + 1
                            if r_info_idx >= len(all_rec_rows): continue
                            
                            r_info = all_rec_rows[r_info_idx]
                            r_info_next = all_rec_rows[r_info_idx + 1] if r_info_idx + 1 < len(all_rec_rows) else []

                            cap = r_info[6] if len(r_info) > 6 else 0
                            gubun = r_info[16] if len(r_info) > 16 else ""
                            res = r_info_next[6] if len(r_info_next) > 6 else 0
                            serial = r_info_next[16] if len(r_info_next) > 16 else serial_no

                            pts = []
                            for offset in range(2, 16):
                                curr_idx = r_info_idx + offset
                                if curr_idx < len(all_rec_rows):
                                    curr_row = all_rec_rows[curr_idx]
                                    nom = curr_row[32] if len(curr_row) > 32 else None
                                    sub = curr_row[36] if len(curr_row) > 36 else None
                                    use_cnt = str(curr_row[37] if len(curr_row) > 37 else "")
                                    ind = curr_row[49] if len(curr_row) > 49 else None
                                    
                                    v60_check = str(curr_row[59] if len(curr_row) > 59 else "").strip()
                                    if v60_check.startswith("(") and v60_check.endswith(")") and v60_check[1:-1].isdigit():
                                        break

                                    if nom is not None and isinstance(nom, (int, float)):
                                        if float(nom) == 0: continue
                                        has_sub = (sub is not None) and (str(sub).strip() != "") and ("회 사용" not in use_cnt)
                                        pts.append({
                                            "nom": nom,
                                            "ind": ind if ind is not None else "-",
                                            "sub": sub if has_sub else "-",
                                        })
                                        if len(pts) >= 10:
                                            break
                            file_instruments.append({
                                "type_kind": "호퍼 저울",
                                "cert_no": cert_no,
                                "device_name": dev_disp,
                                "maker": maker_name if maker_name else "대성계기",
                                "model": "",
                                "serial": serial,
                                "max_str": f"{int(float(cap))} kg" if float(cap).is_integer() else f"{float(cap):.1f} kg",
                                "div_str": format_number(res, get_decimal_places(res)) + " kg",
                                "res_dec": get_decimal_places(res),
                                "gubun": gubun,
                                "pts": pts,
                            })

                # ========================================================
                # [기기 판별 B] 힘 시험기
                # ========================================================
                elif fname.startswith("F") or any(k in raw_device_name for k in ["시험기", "인장", "압축", "만능"]):
                    result_sheets = [s for s in wb_src.sheetnames if s.startswith("교정결과") and "영문" not in s]

                    if result_sheets:
                        ws_res0 = wb_src[result_sheets[0]]
                        for r_idx, row in enumerate(ws_res0.iter_rows(values_only=True)):
                            if r_idx > 15: break
                            row_str = " ".join([str(c or "") for c in row if c is not None])
                            for c in row:
                                if c and re.search(r'[A-Za-z]\d{2}-\d+-\d+', str(c)):
                                    cert_no = str(c).strip()
                            if ("기기명" in row_str or "기      기" in row_str) and not raw_device_name:
                                for c in row:
                                    if c and any(k in str(c) for k in ["시험기", "인장", "압축"]):
                                        raw_device_name = str(c).strip()
                            if "제작회사 및 형식" in row_str or "제작회사" in row_str:
                                for c in row:
                                    if c and "/" in str(c):
                                        parts = str(c).split("/")
                                        maker_name = parts[0].strip()
                                        model_name = ""
                            if ("기기번호" in row_str or "기   기   번" in row_str) and not serial_no:
                                for c in row:
                                    if c and any(ch.isdigit() for ch in str(c)) and "기기" not in str(c) and "*" not in str(c):
                                        serial_no = str(c).strip()

                    if not raw_device_name:
                        raw_device_name = "인장 및 압축 시험기"

                    rec_sheet_names = [s for s in wb_src.sheetnames if "기록지" in s and "영문" not in s]
                    def _sort_rec_key(name):
                        m = re.search(r'\((\d+)\)', name)
                        return int(m.group(1)) if m else 0
                    rec_sheet_names.sort(key=_sort_rec_key)

                    rec_stage_params = []
                    for rs_name in rec_sheet_names:
                        ws_rec_stage = wb_src[rs_name]
                        cap_val = None
                        cap_unit = "kN"
                        div_val = None
                        div_unit = "kN"

                        for r_i, r_vals in enumerate(ws_rec_stage.iter_rows(values_only=True), 1):
                            if r_i > 30:
                                break
                            row_t = " ".join([str(c or "") for c in r_vals if c is not None])
                            if "최대용량" in row_t:
                                for c_i, v in enumerate(r_vals):
                                    if v is not None and isinstance(v, (int, float)) and v > 0:
                                        cap_val = v
                                        for u_c in r_vals[c_i+1:c_i+8]:
                                            if u_c and any(u in str(u_c) for u in ["kN", "N", "tf", "kgf"]):
                                                cap_unit = str(u_c).replace("(", "").replace(")", "").strip()
                                                break
                                        break
                            if "분해능" in row_t or "분 해 능" in row_t:
                                for c_i, v in enumerate(r_vals):
                                    if v is not None and isinstance(v, (int, float)) and v > 0:
                                        div_val = v
                                        for u_c in r_vals[c_i+1:c_i+8]:
                                            if u_c and any(u in str(u_c) for u in ["kN", "N", "tf", "kgf"]):
                                                div_unit = str(u_c).replace("(", "").replace(")", "").strip()
                                                break
                                        break

                        cap_str_rec = f"{int(cap_val) if float(cap_val).is_integer() else cap_val} {cap_unit}" if cap_val is not None else ""
                        div_str_rec = f"{div_val} {div_unit}" if div_val is not None else ""
                        rec_stage_params.append({
                            "max_cap": cap_str_rec,
                            "div_str": div_str_rec,
                        })

                    stages = []
                    for sname in result_sheets:
                        ws = wb_src[sname]
                        rows = list(ws.iter_rows(values_only=True))
                        r = 0
                        while r < len(rows):
                            row_str = " ".join([str(c or "") for c in rows[r]])
                            if any(m in row_str for m in ["압 축 교 정", "압축교정", "인 장 교 정", "인장교정"]):
                                method = "인장" if ("인장" in row_str or "인 장" in row_str) else "압축"
                                start_data_r = -1
                                unit_str = "kN"

                                for dr in range(r + 1, min(r + 6, len(rows))):
                                    dr_str = " ".join([str(c or "") for c in rows[dr]])
                                    if "지시하중" in dr_str:
                                        if dr + 1 < len(rows):
                                            u_str = " ".join([str(c or "") for c in rows[dr + 1]])
                                            m_u = re.search(r'\(([A-Za-z]+)\)', u_str)
                                            if m_u: unit_str = m_u.group(1)
                                        start_data_r = dr + 2
                                        break

                                if start_data_r != -1:
                                    pts = []
                                    curr_r = start_data_r
                                    res_val_str = ""
                                    while curr_r < len(rows):
                                        crow = rows[curr_r]
                                        crow_str = " ".join([str(c or "") for c in crow])
                                        if any(k in crow_str for k in ["상대영점", "분해능", "상대측정불확도", "압 축 교 정", "인 장 교 정", "KCS-QP"]):
                                            break
                                        nom = crow[3] if len(crow) > 3 else None
                                        err = crow[28] if len(crow) > 28 else None
                                        if nom is not None:
                                            try:
                                                f_nom = float(str(nom).replace(",", "").strip())
                                                if f_nom > 0:
                                                    err_str = "-"
                                                    if err is not None and str(err).strip() != "-":
                                                        f_err = float(err)
                                                        err_str = f"{f_err:+.2f} %"
                                                    pts.append({
                                                        "nom": f_nom,
                                                        "ind": err_str,
                                                        "sub": "-"
                                                    })
                                            except Exception:
                                                pass
                                        curr_r += 1

                                    for sr in range(curr_r, min(curr_r + 6, len(rows))):
                                        sr_str = " ".join([str(c or "") for c in rows[sr]])
                                        m_res = re.search(r'분해능\s*:\s*([\d\.]+\s*[A-Za-z]*)', sr_str)
                                        if m_res:
                                            res_val_str = m_res.group(1).strip()
                                            if not any(ch.isalpha() for ch in res_val_str):
                                                res_val_str = f"{res_val_str} {unit_str}"
                                            break

                                    max_nom = max([p["nom"] for p in pts]) if pts else 0
                                    stages.append({
                                        "method": method,
                                        "unit": unit_str,
                                        "fallback_max": f"{int(max_nom) if float(max_nom).is_integer() else max_nom} {unit_str}",
                                        "fallback_div": res_val_str if res_val_str else f"0.01 {unit_str}",
                                        "pts": pts
                                    })
                                    r = curr_r
                                    continue
                            r += 1

                    total_stages = len(stages)
                    for s_idx, stg in enumerate(stages):
                        if total_stages > 1:
                            if s_idx == 0:
                                dev_disp = f"{raw_device_name}\n<{s_idx + 1}단/{total_stages}단>"
                            else:
                                dev_disp = ""
                        else:
                            dev_disp = raw_device_name

                        if s_idx < len(rec_stage_params) and rec_stage_params[s_idx]["max_cap"]:
                            final_max_str = rec_stage_params[s_idx]["max_cap"]
                        else:
                            final_max_str = stg["fallback_max"]

                        if s_idx < len(rec_stage_params) and rec_stage_params[s_idx]["div_str"]:
                            final_div_str = rec_stage_params[s_idx]["div_str"]
                        else:
                            final_div_str = stg["fallback_div"]

                        file_instruments.append({
                            "type_kind": "힘 시험기",
                            "cert_no": cert_no,
                            "device_name": dev_disp,
                            "maker": maker_name,
                            "model": "",
                            "serial": serial_no,
                            "max_str": final_max_str,
                            "div_str": final_div_str,
                            "res_dec": get_decimal_places(stg["pts"][0]["nom"]) if stg["pts"] else 0,
                            "gubun": stg["method"],
                            "pts": stg["pts"],
                            "is_sub_channel": s_idx > 0
                        })

                # ========================================================
                # [기기 판별 C] 온도 기록계 / 온도 지시계
                # ========================================================
                elif fname.startswith("T") or any(k in raw_device_name for k in ["온도", "기록계", "조절계"]):
                    res_sheets = [s for s in wb_src.sheetnames if "교정결과" in s and "영문" not in s]
                    def _sort_sheet_key(s):
                        m = re.search(r'\((\d+)\)', s)
                        return int(m.group(1)) if m else 1
                    res_sheets.sort(key=_sort_sheet_key)

                    if res_sheets:
                        ws_r0 = wb_src[res_sheets[0]]
                        for r_idx, row in enumerate(ws_r0.iter_rows(values_only=True)):
                            if r_idx > 18: break
                            c3 = str(row[2] if len(row) > 2 else "").replace(" ", "")
                            c11 = str(row[10] if len(row) > 10 else "").strip()
                            if "기기명" in c3 and not raw_device_name:
                                raw_device_name = c11
                            if "제작회사" in c3 and not maker_name:
                                parts = c11.split("/")
                                maker_name = parts[0].strip()
                                if len(parts) > 1 and not model_name:
                                    model_name = parts[1].strip()
                            if "기기번호" in c3 and not serial_no:
                                serial_no = c11

                    if not raw_device_name:
                        raw_device_name = "온도 기록계(센서제외)"

                    temp_channels = []

                    for sname in res_sheets:
                        ws_temp = wb_src[sname]
                        rows_temp = list(ws_temp.iter_rows(values_only=True))

                        sheet_sensor_type = "K"
                        for r_idx in range(min(22, len(rows_temp))):
                            r_str = " ".join([str(v or "") for v in rows_temp[r_idx]])
                            if "온도교정" in r_str or "온도 교정" in r_str:
                                m_t = re.search(r'온도\s*교정\s*[\(（](.*)', r_str)
                                if m_t:
                                    t_val = m_t.group(1)
                                    if '(단위' in t_val:
                                        t_val = t_val.split('(단위')[0]
                                    t_val = t_val.strip()
                                    if t_val.endswith(')') or t_val.endswith('）'):
                                        t_val = t_val[:-1].strip()
                                    sheet_sensor_type = re.sub(r'\s+[tT]ype', '', t_val).strip()
                                    break

                        r_i = 0
                        while r_i < len(rows_temp):
                            r_vals = rows_temp[r_i]
                            r_str = " ".join([str(v or "") for v in r_vals if v is not None])

                            m_ch = re.search(r'[\(（]\s*(CH\s*\d+)\s*[\)）]', r_str, re.IGNORECASE)
                            if m_ch:
                                ch_label = m_ch.group(1).upper().replace(" ", "")

                                indiv_serial = ""
                                m_sn = re.search(r'기기번호\s*:\s*([^|\(,\n]+)', r_str)
                                if m_sn:
                                    indiv_serial = m_sn.group(1).strip()
                                    indiv_serial = re.sub(r'\(단위.*?\)', '', indiv_serial).strip()

                                ch_type = sheet_sensor_type
                                m_local_t = re.search(r'온도\s*교정\s*[\(（](.*)', r_str)
                                if m_local_t:
                                    lt_val = m_local_t.group(1)
                                    if '(단위' in lt_val:
                                        lt_val = lt_val.split('(단위')[0]
                                    lt_val = lt_val.strip()
                                    if lt_val.endswith(')') or lt_val.endswith('）'):
                                        lt_val = lt_val[:-1].strip()
                                    ch_type = re.sub(r'\s+[tT]ype', '', lt_val).strip()

                                nom_col = None
                                ind_col = None
                                data_start = None

                                for dr in range(r_i + 1, min(r_i + 5, len(rows_temp))):
                                    d_row = rows_temp[dr]
                                    d_str = " ".join([str(v or "") for v in d_row if v is not None])
                                    if "기준온도" in d_str and "지시온도" in d_str:
                                        for c_idx, v in enumerate(d_row):
                                            if v and "기준온도" in str(v): nom_col = c_idx
                                            if v and "지시온도" in str(v): ind_col = c_idx
                                        data_start = dr + 1
                                        if data_start < len(rows_temp):
                                            sub_str = " ".join([str(v or "") for v in rows_temp[data_start] if v is not None])
                                            if "기준온도-지시온도" in sub_str or "신뢰수준" in sub_str or "℃" in sub_str:
                                                data_start += 1
                                        break

                                pts = []
                                if data_start and nom_col is not None and ind_col is not None:
                                    curr = data_start
                                    while curr < len(rows_temp):
                                        crow = rows_temp[curr]
                                        c_str = " ".join([str(v or "") for v in crow if v is not None])
                                        if any(k in c_str for k in ["끝.", "※", "권장차기교정일", "KCS-QP"]):
                                            break
                                        if re.search(r'[\(（]\s*CH\s*\d+\s*[\)）]', c_str, re.IGNORECASE):
                                            break
                                        if "온도교정" in c_str or "기기명" in c_str:
                                            break

                                        v_nom = crow[nom_col] if len(crow) > nom_col else None
                                        v_ind = crow[ind_col] if len(crow) > ind_col else None

                                        if v_nom is not None:
                                            try:
                                                clean_nom = str(v_nom).replace(",", "").strip()
                                                clean_ind = str(v_ind).replace(",", "").strip() if v_ind is not None else None
                                                f_nom = float(clean_nom)
                                                f_ind = float(clean_ind) if clean_ind is not None else None
                                                pts.append({
                                                    "nom": f_nom,
                                                    "ind": f_ind if f_ind is not None else "-",
                                                    "ind_raw": v_ind,
                                                    "sub": "-"
                                                })
                                                if len(pts) >= 10:
                                                    break
                                            except Exception:
                                                pass
                                        curr += 1

                                temp_channels.append({
                                    "ch_label": ch_label,
                                    "indiv_serial": indiv_serial,
                                    "sensor_type": ch_type,
                                    "pts": pts
                                })
                            r_i += 1

                    if not temp_channels and res_sheets:
                        ws_res_single = wb_src[res_sheets[0]]
                        res_rows = list(ws_res_single.iter_rows(values_only=True))
                        fallback_type = "K"
                        for r in range(10, min(17, len(res_rows))):
                            val_t = str(res_rows[r][1] if len(res_rows[r]) > 1 else "") + " " + str(res_rows[r][2] if len(res_rows[r]) > 2 else "")
                            if "온도교정" in val_t or "온도 교정" in val_t:
                                m_s = re.search(r'온도\s*교정\s*[\(（](.*)', val_t)
                                if m_s:
                                    lt_val = m_s.group(1)
                                    if '(단위' in lt_val:
                                        lt_val = lt_val.split('(단위')[0]
                                    lt_val = lt_val.strip()
                                    if lt_val.endswith(')') or lt_val.endswith('）'):
                                        lt_val = lt_val[:-1].strip()
                                    fallback_type = re.sub(r'\s+[tT]ype', '', lt_val).strip()
                                break

                        pts = []
                        for r in range(17, min(35, len(res_rows))):
                            c3_v = res_rows[r][2] if len(res_rows[r]) > 2 else None
                            c11_v = res_rows[r][10] if len(res_rows[r]) > 10 else None
                            if c3_v is not None and isinstance(c3_v, (int, float)):
                                pts.append({
                                    "nom": c3_v,
                                    "ind": c11_v if c11_v is not None else "-",
                                    "ind_raw": c11_v,
                                    "sub": "-",
                                })
                                if len(pts) >= 10:
                                    break
                            elif str(c3_v).strip() == "끝.":
                                break

                        temp_channels.append({
                            "ch_label": "CH1",
                            "indiv_serial": "",
                            "sensor_type": fallback_type,
                            "pts": pts
                        })

                    total_ch = len(temp_channels)

                    for ch_idx, ch in enumerate(temp_channels):
                        res_dec = 0
                        for p in ch["pts"]:
                            raw_ind = p.get("ind_raw", None)
                            if raw_ind is not None:
                                d = get_decimal_places(raw_ind)
                                if d > res_dec:
                                    res_dec = d
                        div_str = f"0.{'0'*(res_dec-1)}1 ℃" if res_dec > 0 else "1 ℃"

                        clean_base_dev = raw_device_name.replace("온도기록계", "온도 기록계").replace("온도지시계", "온도 지시계")
                        max_channel_str = f"<{ch['ch_label']}/{total_ch}CH>" if total_ch > 1 else ""

                        final_serial = ch["indiv_serial"] if ch["indiv_serial"] else serial_no
                        if not final_serial or final_serial == "교정결과 참조":
                            final_serial = serial_no if serial_no != "교정결과 참조" else ""

                        file_instruments.append({
                            "type_kind": "온도기록계",
                            "cert_no": cert_no,
                            "device_name": clean_base_dev,
                            "maker": maker_name,
                            "model": model_name,
                            "serial": final_serial,
                            "max_str": max_channel_str,
                            "div_str": div_str,
                            "res_dec": res_dec,
                            "gubun": ch["sensor_type"],
                            "pts": ch["pts"],
                            "is_sub_channel": ch_idx > 0
                        })

                # ========================================================
                # [기기 판별 D] 압력계 / 차압계
                # ========================================================
                elif fname.startswith("P") or any(k in raw_device_name for k in ["압력", "진공", "차압"]) or any(u in capacity_raw for u in ["bar", "MPa", "kPa", "mmH2O"]):
                    result_sheet_names = [s for s in wb_src.sheetnames if s.startswith("교정결과") and "영문" not in s]
                    ws_res = wb_src[result_sheet_names[0]] if result_sheet_names else wb_src.active

                    p_dev_name = raw_device_name
                    p_maker = maker_name
                    p_capacity = capacity_raw
                    p_serial = serial_no

                    for row in ws_res.iter_rows(values_only=True):
                        for c in row:
                            if not c: continue
                            cs = str(c).strip()
                            if ("기기명" in cs or "기   기   명" in cs) and ":" in cs and not p_dev_name:
                                p_dev_name = cs.split(":", 1)[1].strip()
                            elif ("제작회사" in cs or "제 작 회 사" in cs) and ":" in cs and not p_maker:
                                p_maker = cs.split(":", 1)[1].strip()
                            elif ("형식" in cs or "형   식" in cs) and ":" in cs:
                                val_f = cs.split(":", 1)[1].strip()
                                if val_f: p_capacity = val_f
                            elif ("기기번호" in cs or "기   기   번   호" in cs) and ":" in cs and not p_serial:
                                p_serial = cs.split(":", 1)[1].strip()

                    if not p_dev_name:
                        p_dev_name = "차압계" if "차압" in capacity_raw else "게이지압용 압력계(다이얼형)"

                    p_div_str = ""
                    if "기록지" in wb_src.sheetnames:
                        ws_rec_p = wb_src["기록지"]
                        rows_rec_p = list(ws_rec_p.iter_rows(values_only=True))
                        min_scale = None
                        resolution = None
                        p_unit = ""

                        for r_i, r_vals in enumerate(rows_rec_p):
                            row_text = " ".join([str(v or "") for v in r_vals])
                            if "최소 눈금" in row_text or "최소눈금" in row_text:
                                if r_i + 1 < len(rows_rec_p):
                                    next_r = rows_rec_p[r_i + 1]
                                    for c_i, v in enumerate(r_vals):
                                        if v and ("최소 눈금" in str(v) or "최소눈금" in str(v)):
                                            val_m = next_r[c_i]
                                            if val_m is not None and isinstance(val_m, (int, float)):
                                                min_scale = val_m
                                                for u_val in next_r[c_i+1:c_i+12]:
                                                    if u_val and any(u in str(u_val) for u in ["bar", "MPa", "kPa", "Pa", "psi", "cmHg", "mmHg", "kgf", "mmH2O"]):
                                                        p_unit = str(u_val).replace("(", "").replace(")", "").strip()
                                                        break
                                        if v and "분해능" in str(v):
                                            val_d = next_r[c_i]
                                            if val_d is not None and isinstance(val_d, (int, float)):
                                                resolution = val_d
                                                if not p_unit:
                                                    for u_val in next_r[c_i+1:c_i+12]:
                                                        if u_val and any(u in str(u_val) for u in ["bar", "MPa", "kPa", "Pa", "psi", "cmHg", "mmHg", "kgf", "mmH2O"]):
                                                            p_unit = str(u_val).replace("(", "").replace(")", "").strip()
                                                            break
                                    if min_scale is not None and resolution is not None:
                                        break

                        if min_scale is not None and resolution is not None and resolution > 0:
                            ratio = min_scale / resolution
                            r_str = f"{int(ratio) if float(ratio).is_integer() else ratio}"
                            m_str = f"{int(min_scale) if float(min_scale).is_integer() else min_scale}"
                            p_div_str = f"{m_str}({r_str}) {p_unit}".strip()

                    pts = []
                    is_diff_pressure = ("차압" in p_dev_name) or ("차압" in raw_device_name)

                    if is_diff_pressure and "기록지" in wb_src.sheetnames:
                        ws_rec_diff = wb_src["기록지"]
                        rows_rec_all = list(ws_rec_diff.iter_rows(values_only=True))

                        col_1st = None
                        row_1st_header = None
                        press_unit = ""

                        for r_idx, r_vals in enumerate(rows_rec_all):
                            for c_idx, val in enumerate(r_vals):
                                if val and str(val).strip() == "1회":
                                    col_1st = c_idx
                                    row_1st_header = r_idx
                                    break
                            if col_1st is not None:
                                break

                        if col_1st is not None and row_1st_header is not None:
                            unit_row = rows_rec_all[row_1st_header + 1] if row_1st_header + 1 < len(rows_rec_all) else []
                            col_nom = None

                            for c_back in range(col_1st - 1, -1, -1):
                                if c_back < len(unit_row) and unit_row[c_back] is not None and str(unit_row[c_back]).strip() != "":
                                    col_nom = c_back
                                    press_unit = str(unit_row[c_back]).replace("(", "").replace(")", "").strip()
                                    break

                            if col_nom is None:
                                col_nom = col_1st - 1

                            if not press_unit and col_1st < len(unit_row) and unit_row[col_1st]:
                                press_unit = str(unit_row[col_1st]).replace("(", "").replace(")", "").strip()

                            col_pt = 1
                            for c_check in range(0, col_nom):
                                c_header = str(rows_rec_all[row_1st_header - 1][c_check] if row_1st_header - 1 >= 0 else "")
                                if "측정점" in c_header:
                                    col_pt = c_check
                                    break

                            start_data_r = -1
                            for r_scan in range(row_1st_header + 1, min(row_1st_header + 6, len(rows_rec_all))):
                                if ws_rec_diff.row_dimensions[r_scan + 1].hidden:
                                    continue
                                v_test = rows_rec_all[r_scan][col_1st] if len(rows_rec_all[r_scan]) > col_1st else None
                                try:
                                    if v_test is not None:
                                        float(str(v_test).replace(" ", "").replace(",", ""))
                                        start_data_r = r_scan
                                        break
                                except ValueError:
                                    pass

                            raw_points = []
                            if start_data_r != -1:
                                for r_data in range(start_data_r, len(rows_rec_all)):
                                    if ws_rec_diff.row_dimensions[r_data + 1].hidden:
                                        continue

                                    v_nom_cell = ws_rec_diff.cell(row=r_data+1, column=col_nom+1)
                                    v_ind_cell = ws_rec_diff.cell(row=r_data+1, column=col_1st+1)
                                    
                                    raw_nom = format_meaningful_decimals(v_nom_cell.value)
                                    raw_ind = get_formatted_value(v_ind_cell)
                                    
                                    v_pt = rows_rec_all[r_data][col_pt] if len(rows_rec_all[r_data]) > col_pt else None

                                    if raw_nom == "-" and raw_ind is None:
                                        break

                                    try:
                                        f_nom = float(str(v_nom_cell.value).replace(" ", "").replace(",", ""))
                                        f_ind = float(str(v_ind_cell.value).replace(" ", "").replace(",", "")) if raw_ind is not None else None
                                        
                                        try:
                                            pt_num = int(float(str(v_pt).strip())) if v_pt is not None else len(raw_points) + 1
                                        except Exception:
                                            pt_num = len(raw_points) + 1

                                        raw_points.append({
                                            "pt": pt_num,
                                            "nom_f": f_nom,
                                            "ind_f": f_ind,
                                            "raw_nom": raw_nom,
                                            "raw_ind": raw_ind if raw_ind is not None else "-"
                                        })
                                    except Exception:
                                        break

                            if raw_points:
                                peak_idx = 0
                                max_val = -float("inf")
                                for idx, item in enumerate(raw_points):
                                    val = item["nom_f"]
                                    if val >= max_val:
                                        max_val = val
                                        peak_idx = idx
                                    else:
                                        break
                                sliced_points = raw_points[: peak_idx + 1]

                                first_pt = sliced_points[0]
                                if first_pt["pt"] == 1 and abs(first_pt["nom_f"]) < 1e-7:
                                    sliced_points = sliced_points[1:]

                                for p_i, item in enumerate(sliced_points[:10]):
                                    nom_final = f"{item['raw_nom']} {press_unit}" if p_i == 0 and press_unit else item['raw_nom']
                                    ind_final = f"{item['raw_ind']} {press_unit}" if p_i == 0 and press_unit and item['raw_ind'] != "-" else item['raw_ind']

                                    pts.append({
                                        "nom": nom_final,
                                        "ind": ind_final,
                                        "sub": "-",
                                    })

                    else:
                        rows_res_p = list(ws_res.iter_rows(values_only=True))
                        data_start_r = -1
                        press_unit = "MPa"
                        std_col = 15
                        ind_col = 35

                        for r_i, r_vals in enumerate(rows_res_p):
                            row_str = " ".join([str(v or "") for v in r_vals])
                            if ("표준" in row_str and "압력" in row_str and "지시" in row_str) or ("표준기" in row_str and "지시" in row_str):
                                for c_i, v in enumerate(r_vals):
                                    if v and "표준" in str(v) and "압력" in str(v):
                                        std_col = c_i
                                    elif v and "지시" in str(v):
                                        ind_col = c_i

                                for offset in [1, 2, 3]:
                                    if r_i + offset < len(rows_res_p):
                                        unit_row_str = " ".join([str(v or "") for v in rows_res_p[r_i + offset]])
                                        for u in ["MPa", "kPa", "bar", "psi", "cmHg", "mmHg", "Pa", "kgf"]:
                                            if u in unit_row_str:
                                                press_unit = u
                                                break
                                        if press_unit != "MPa":
                                            break

                                for check_r in range(r_i + 1, min(r_i + 6, len(rows_res_p))):
                                    if ws_res.row_dimensions[check_r + 1].hidden:
                                        continue
                                    v_check = rows_res_p[check_r][std_col] if len(rows_res_p[check_r]) > std_col else None
                                    try:
                                        if v_check is not None:
                                            float(str(v_check).replace(",", ""))
                                            data_start_r = check_r
                                            break
                                    except ValueError:
                                        pass
                                break

                        if data_start_r != -1:
                            for r_i in range(data_start_r, min(data_start_r + 25, len(rows_res_p))):
                                if ws_res.row_dimensions[r_i + 1].hidden:
                                    continue
                                
                                crow = rows_res_p[r_i]
                                crow_str = " ".join([str(v or "") for v in crow])
                                if any(k in crow_str for k in ["*", "※", "끝."]):
                                    break
                                
                                std_cell = ws_res.cell(row=r_i+1, column=std_col+1)
                                ind_cell = ws_res.cell(row=r_i+1, column=ind_col+1)
                                
                                if std_cell.value is not None:
                                    try:
                                        f_std = float(str(std_cell.value).replace(",", ""))
                                        if abs(f_std) < 1e-7 and len(pts) == 0:
                                            continue
                                            
                                        std_s = format_meaningful_decimals(std_cell.value)
                                        if std_s != "-":
                                            std_s = f"{std_s} {press_unit}".strip()
                                        
                                        ind_s = get_formatted_value(ind_cell)
                                        if ind_s is not None:
                                            ind_s = f"{ind_s} {press_unit}".strip()
                                        else:
                                            ind_s = "-"

                                        pts.append({
                                            "nom": std_s,
                                            "ind": ind_s,
                                            "sub": "-",
                                        })
                                        if len(pts) >= 10:
                                            break
                                    except Exception:
                                        pass

                    file_instruments.append({
                        "type_kind": "압력계",
                        "cert_no": cert_no,
                        "device_name": p_dev_name.replace("게이지압용 ", "게이지압용\n"),
                        "maker": p_maker,
                        "model": "",
                        "serial": p_serial,
                        "max_str": p_capacity,
                        "div_str": p_div_str,
                        "res_dec": 0,
                        "gubun": "",
                        "pts": pts,
                    })

                # ========================================================
                # [기기 판별 E] 전기식 지시 저울
                # ========================================================
                else:
                    ws_res = wb_src["교정결과"]
                    rows = list(ws_res.iter_rows(values_only=True))
                    gubun_val = "크레인" if "크레인" in raw_device_name else ""

                    for r in range(0, min(20, len(rows))):
                        c4 = str(rows[r][3] if len(rows[r]) > 3 else "").replace(" ", "")
                        c10 = str(rows[r][9] if len(rows[r]) > 9 else "").strip()
                        if "제작회사" in c4 and not maker_name:
                            maker_name = c10.split("/")[0].strip()
                        if "기기번호" in c4 and not serial_no:
                            serial_no = c10
                        if "최대용량" in c4 and not capacity_raw:
                            capacity_raw = c10
                        if "분해능" in c4 and not scale_raw:
                            scale_raw = c10

                    max_str, div_str, is_dual = format_capacity_and_resolution_robust(capacity_raw, scale_raw)
                    res_dec = determine_scale_decimal_places(div_str, max_str, is_dual)

                    pts = []
                    start_r = 0
                    for r in range(0, min(25, len(rows))):
                        if "저울지시값" in str(rows[r][2] if len(rows[r]) > 2 else ""):
                            start_r = r + 1
                            break

                    if start_r > 0:
                        for r in range(start_r, min(start_r + 20, len(rows))):
                            ind_val = rows[r][2] if len(rows[r]) > 2 else None
                            std_val = rows[r][8] if len(rows[r]) > 8 else None

                            val_str = str(rows[r][2] if len(rows[r]) > 2 else "")
                            if "※" in val_str or "끝" in val_str:
                                break
                            
                            if ind_val is None and std_val is None:
                                continue

                            try:
                                if float(ind_val) == 0: continue
                            except (ValueError, TypeError):
                                pass

                            try:
                                nom_val = round(float(std_val)) if std_val is not None and isinstance(std_val, (int, float)) else ind_val
                                if nom_val is None or float(nom_val) == 0: 
                                    continue
                            except (ValueError, TypeError):
                                continue

                            pts.append({
                                "nom": nom_val,
                                "ind": ind_val,
                                "sub": "-",
                            })
                            if len(pts) >= 10:
                                break

                    dev_disp = "전기식 지시 저울\n(크레인)" if "크레인" in raw_device_name else raw_device_name.replace("전기식지시저울", "전기식 지시 저울")

                    file_instruments.append({
                        "type_kind": "전기식 저울",
                        "cert_no": cert_no,
                        "device_name": dev_disp,
                        "maker": maker_name,
                        "model": "",
                        "serial": serial_no,
                        "max_str": max_str,
                        "div_str": div_str,
                        "res_dec": res_dec,
                        "gubun": gubun_val,
                        "pts": pts,
                    })

                wb_src.close()

                for sub_i, inst in enumerate(file_instruments):
                    inst["cert_file_no"] = file_idx if sub_i == 0 else ""
                    all_instruments.append(inst)

            total_slots_needed = sum([2 if len(inst["pts"]) > 5 else 1 for inst in all_instruments])
            total_pages_needed = max(1, (total_slots_needed + 9) // 10)

            wb_out = openpyxl.load_workbook(TEMPLATE_FILE)
            ws_template = wb_out["교정점검표"]

            sheet_list = [ws_template]
            for p in range(2, total_pages_needed + 1):
                ws_new = wb_out.copy_worksheet(ws_template)
                ws_new.title = f"교정점검표 ({p})"
                sheet_list.append(ws_new)

            for ws_target in sheet_list:
                if global_client_name:
                    ws_target["C1"] = str(global_client_name)
                    ws_target["C1"].alignment = Alignment(horizontal="left", vertical="center")
                
                # 10칸 기록지 우측 상단 헤더 좌표 세팅
                ws_target["H1"] = "-"
                ws_target["H1"].alignment = Alignment(horizontal="center", vertical="center")
                ws_target["H2"] = "-"
                ws_target["H2"].alignment = Alignment(horizontal="center", vertical="center")
                ws_target["I2"] = "   月    日"
                ws_target["I2"].alignment = Alignment(horizontal="center", vertical="center")

                # 빈칸 초기화
                for slot_i in range(10):
                    ws_target.cell(row=4 + slot_i * 5, column=1, value="")

            current_global_slot = 0
            for inst in all_instruments:
                pts = inst["pts"]
                res_dec = inst["res_dec"]
                slots_for_inst = 2 if len(pts) > 5 else 1

                page_idx = current_global_slot // 10
                slot_in_page = current_global_slot % 10
                ws_curr = sheet_list[page_idx]

                start_row = 4 + slot_in_page * 5

                # A열: No
                if not del_number:
                    if inst["cert_file_no"]:
                        ws_curr.cell(row=start_row, column=1, value=inst["cert_file_no"])

                # B열: 기기명
                if inst.get("is_sub_channel"):
                    dev_name_to_write = ""
                else:
                    dev_name_to_write = str(inst["device_name"])
                cell_dev = ws_curr.cell(row=start_row, column=2, value=dev_name_to_write)
                cell_dev.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

                # C열: 제조사 / 모델
                if inst.get("is_sub_channel"):
                    maker_to_write = ""
                    model_to_write = ""
                else:
                    maker_to_write = str(inst["maker"])
                    model_to_write = "" if inst["type_kind"] in ["압력계", "힘 시험기"] else str(inst["model"])
                    
                ws_curr.cell(row=start_row, column=3, value=maker_to_write)
                ws_curr.cell(row=start_row + 2, column=3, value=model_to_write)

                # D열: 기기번호
                ws_curr.cell(row=start_row, column=4, value=str(inst["serial"]))

                # E열: MAX / DIV
                ws_curr.cell(row=start_row, column=5, value=str(inst["max_str"]))
                ws_curr.cell(row=start_row + 2, column=5, value=str(inst["div_str"]))

                # F열: 방식 / 구분 / 타입
                ws_curr.cell(row=start_row, column=6, value=str(inst["gubun"]))

                # G열 & H열: 범위 및 측정결과
                for p_idx, pt in enumerate(pts[:10]):
                    target_r = start_row + p_idx
                    if target_r > 65:
                        break

                    if not del_nominal:
                        if isinstance(pt["nom"], str):
                            ws_curr.cell(row=target_r, column=7, value=pt["nom"])
                        else:
                            nom_str = format_number(pt["nom"], res_dec)
                            ws_curr.cell(row=target_r, column=7, value=nom_str)

                    if not del_data:
                        if isinstance(pt["ind"], str):
                            ws_curr.cell(row=target_r, column=8, value=pt["ind"])
                        else:
                            ind_str = format_number(pt["ind"], res_dec)
                            sub_val = pt.get("sub", "-")
                            if sub_val != "-" and sub_val != 0 and sub_val != "0":
                                sub_str = format_sub_weight(sub_val)
                                ws_curr.cell(row=target_r, column=8, value=f"{ind_str} (가:{sub_str})")
                            else:
                                ws_curr.cell(row=target_r, column=8, value=ind_str)

                # I열: 교정번호
                ws_curr.cell(row=start_row, column=9, value="")

                current_global_slot += slots_for_inst

            output_buffer = io.BytesIO()
            wb_out.save(output_buffer)
            output_buffer.seek(0)

            safe_client = re.sub(r'[\/:*?"<>|() ]', "", str(global_client_name))
            out_name = f"기록지_{safe_client}_총{len(uploaded_files)}건_({total_pages_needed}P).xlsx"

            # [결과 출력] 전체 너비 컨테이너에서 렌더링
            with bottom_result_container:
                st.markdown("<div style='margin-top: 2.2rem; border-top: 1px solid #E2E8F0; padding-top: 1.8rem;'></div>", unsafe_allow_html=True)
                
                # 다운로드 버튼 및 완료 배지
                col_d1, col_d2 = st.columns([3.5, 1.2], gap="medium")
                with col_d1:
                    st.download_button(
                        label="👉 [딸깍] 기록지 다운로드",
                        data=output_buffer,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                    )
                with col_d2:
                    st.success(f"총 {len(uploaded_files)}건 완료 ({total_pages_needed}P)")

                # 결과 메트릭 카드 4개를 전체 너비 가로 1열로 배치
                st.markdown("<div style='margin-top: 1.2rem;'></div>", unsafe_allow_html=True)
                m_c1, m_c2, m_c3, m_c4 = st.columns(4, gap="medium")
                m_c1.metric("🏢 의뢰 업체명", str(global_client_name) if global_client_name else "미확인")
                m_c2.metric("📑 업로드 성적서", f"{len(uploaded_files)} 건")
                m_c3.metric("🎯 총 사용 슬롯", f"{total_slots_needed} 칸")
                m_c4.metric("📄 자동 생성 페이지", f"{total_pages_needed} 페이지")

        except Exception as e:
            st.error(f"분석 중 오류가 발생했습니다: {e}")

# ----------------- 하단 푸터 -----------------
st.markdown("<div style='margin-top: 3.5rem;'></div>", unsafe_allow_html=True)
st.caption("⚡ **KCS 딸깍 기록지** | Crafted by **D** | Automated Calibration Record Generator")
