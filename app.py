import streamlit as st
import pandas as pd
from datetime import datetime, date
import urllib.parse
import os
import re

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

# ---------------------------------------------------------
# 1. 기본 페이지 설정 및 스타일
# ---------------------------------------------------------
st.set_page_config(
    page_title="차량 QR 이력·상태 관리 시스템",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .car-header-banner {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        padding: 22px 28px;
        border-radius: 14px;
        margin-bottom: 18px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    }
    .car-header-title {
        font-size: 26px;
        font-weight: 800;
        margin: 0;
    }
    .car-header-sub {
        font-size: 14px;
        opacity: 0.9;
        margin-top: 8px;
    }
    .status-card {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 14px 16px;
        margin-bottom: 10px;
    }
    .kpi-box {
        background-color: #ffffff;
        border: 1px solid #e5e7eb;
        border-left: 5px solid #2a5298;
        border-radius: 10px;
        padding: 16px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.03);
    }
</style>
""", unsafe_allow_html=True)

EXCEL_FILE = "car_data.xlsx"

# 구글 시트 연동 활성화 스위치 (1단계 앱 화면 확인 후 True로 전환 예정)
ENABLE_GOOGLE_SHEETS = False

# ---------------------------------------------------------
# 2. 초기 기본 데이터 정의 (주차위치, 키위치 제거 버전)
# ---------------------------------------------------------
def get_default_dataframes():
    df_cars = pd.DataFrame([
        {
            "차량명": "본사-1",
            "차종": "아이오닉 6 (전기차)",
            "차량정보": "88하 1024",
            "소속본부": "본사 (지원처)",
            "용도구분": "일반 업무용",
            "연료": "전기(EV)",
            "담당자_연락처": "본사 지원처 (내선 0000)",
            "최근 주행거리": 24150,
            "정기검사 만료일": "2027-03-01",
            "보험만료일": "2027-02-15"
        },
        {
            "차량명": "본사-2",
            "차종": "쏘나타 DN8",
            "차량정보": "123가 4567",
            "소속본부": "본사 (지원처)",
            "용도구분": "일반 업무용",
            "연료": "가솔린",
            "담당자_연락처": "본사 지원처 (내선 0000)",
            "최근 주행거리": 48200,
            "정기검사 만료일": "2026-12-15",
            "보험만료일": "2027-01-20"
        },
        {
            "차량명": "북부-1",
            "차종": "스포티지 하이브리드",
            "차량정보": "45나 7890",
            "소속본부": "수도권북부본부",
            "용도구분": "현장 점검용",
            "연료": "하이브리드",
            "담당자_연락처": "수도권북부 운영팀",
            "최근 주행거리": 61500,
            "정기검사 만료일": "2026-11-20",
            "보험만료일": "2026-12-01"
        },
        {
            "차량명": "영남-1",
            "차종": "카니발 디젤",
            "차량정보": "78다 1122",
            "소속본부": "영남본부",
            "용도구분": "현장 점검용",
            "연료": "디젤",
            "담당자_연락처": "영남본부 운영팀",
            "최근 주행거리": 89400,
            "정기검사 만료일": "2027-04-10",
            "보험만료일": "2027-03-25"
        }
    ])

    df_rules = pd.DataFrame([
        {"항목": "엔진오일", "권장주기_km": 10000, "권장주기_개월": 12},
        {"항목": "에어컨필터", "권장주기_km": 15000, "권장주기_개월": 6},
        {"항목": "브레이크패드", "권장주기_km": 40000, "권장주기_개월": 24},
        {"항목": "타이어 위치교환/교체", "권장주기_km": 20000, "권장주기_개월": 24},
        {"항목": "와이퍼 블레이드", "권장주기_km": 10000, "권장주기_개월": 12},
        {"항목": "12V 보조배터리", "권장주기_km": 50000, "권장주기_개월": 36}
    ])

    df_cons = pd.DataFrame([
        {"차량명": "본사-1", "항목": "에어컨필터", "교체주기(Km)": 15000, "교체주기(개월)": 6, "마지막 교체일": "2026-06-15", "마지막 교체 주행거리": 18000},
        {"차량명": "본사-1", "항목": "타이어 위치교환/교체", "교체주기(Km)": 20000, "교체주기(개월)": 24, "마지막 교체일": "2025-11-10", "마지막 교체 주행거리": 10000},
        {"차량명": "본사-2", "항목": "엔진오일", "교체주기(Km)": 10000, "교체주기(개월)": 12, "마지막 교체일": "2026-03-10", "마지막 교체 주행거리": 40000},
        {"차량명": "본사-2", "항목": "에어컨필터", "교체주기(Km)": 15000, "교체주기(개월)": 6, "마지막 교체일": "2026-03-10", "마지막 교체 주행거리": 40000},
        {"차량명": "북부-1", "항목": "엔진오일", "교체주기(Km)": 10000, "교체주기(개월)": 12, "마지막 교체일": "2025-12-05", "마지막 교체 주행거리": 50500},
        {"차량명": "영남-1", "항목": "엔진오일", "교체주기(Km)": 10000, "교체주기(개월)": 12, "마지막 교체일": "2026-05-20", "마지막 교체 주행거리": 82000}
    ])

    df_reps = pd.DataFrame([
        {"차량명": "본사-1", "일자": "2026-07-20", "주행거리(Km)": 19500, "정비구분": "소모품", "내용": "에어컨필터 및 와이퍼 교체", "정비소": "블루핸즈 분당점"},
        {"차량명": "본사-2", "일자": "2026-03-10", "주행거리(Km)": 40000, "정비구분": "소모품", "내용": "엔진오일 세트 교체", "정비소": "블루핸즈 수지점"},
        {"차량명": "북부-1", "일자": "2026-08-12", "주행거리(Km)": 58000, "정비구분": "일반정비", "내용": "타이어 펑크 패치 수리", "정비소": "오토큐 의정부점"}
    ])

    df_fbs = pd.DataFrame([
        {"차량명": "본사-1", "날짜": "2026-10-22", "작성자": "김주임", "확인일자": "2026-10-22", "분류": "기타", "내용": "워셔액 보충 완료 및 실내 점검 완료", "처리상태": "🟢 조치완료"},
        {"차량명": "북부-1", "날짜": "2026-10-24", "작성자": "이대리", "확인일자": "-", "분류": "소음/진동", "내용": "방지턱 넘을 때 하부 소음 약간 발생", "처리상태": "🔴 확인필요"}
    ])

    df_man = pd.DataFrame([
        {"순서": 1, "단계명": "1단계: 비상등 점등 및 안전 확보", "행동요령": "사고 즉시 비상등을 켜고, 차량 이동이 가능하면 갓길 등 안전지대로 대피합니다."},
        {"순서": 2, "단계명": "2단계: 부상자 구호 및 119/112 신고", "행동요령": "동승자 및 상대방 부상 여부를 최우선 확인하고 필요시 119 및 경찰(112)에 신고합니다."},
        {"순서": 3, "단계명": "3단계: 현장 사진 촬영 및 증거 확보", "행동요령": "원거리(차선 포함), 근거리(파손 부위), 상대 차량 번호판 및 블랙박스 영상을 확보합니다."},
        {"순서": 4, "단계명": "4단계: 보험사 접수 및 담당 부서 보고", "행동요령": "공용차량 전담 보험사 콜센터에 사고를 접수하고 소속 본부 차량 담당자에게 즉시 보고합니다."}
    ])

    return df_cars, df_rules, df_cons, df_reps, df_fbs, df_man


# ---------------------------------------------------------
# 3. 컬럼 무결성 보장 및 데이터 로드/저장 함수
# ---------------------------------------------------------
def normalize_dataframe(df, default_df, drop_legacy_cols=True):
    if df is None or df.empty:
        return default_df.copy()
    df_clean = df.copy()
    # 주차위치, 키위치 등 이전 버전 잔재 컬럼 자동 제거
    if drop_legacy_cols:
        for legacy_col in ["주차위치", "주차 위치", "키위치", "키 위치", "보관장소"]:
            if legacy_col in df_clean.columns:
                df_clean = df_clean.drop(columns=[legacy_col])
    # 필수 컬럼 누락 시 기본값으로 보충 (KeyError 방지)
    for col in default_df.columns:
        if col not in df_clean.columns:
            df_clean[col] = default_df[col].iloc[0] if len(default_df) > 0 else ""
    return df_clean[default_df.columns]


def get_gspread_client():
    if not GSPREAD_AVAILABLE or not ENABLE_GOOGLE_SHEETS:
        return None
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            if "private_key" in creds_dict:
                raw_key = str(creds_dict["private_key"])
                cleaned = raw_key.replace("-----BEGIN PRIVATE KEY-----", "").replace("-----END PRIVATE KEY-----", "")
                cleaned = cleaned.replace("\\n", "\n").replace("\r", "")
                pure_chars = "".join(re.findall(r'[A-Za-z0-9+/=]', cleaned))
                missing_padding = len(pure_chars) % 4
                if missing_padding:
                    pure_chars += "=" * (4 - missing_padding)
                chunks = [pure_chars[i:i+64] for i in range(0, len(pure_chars), 64)]
                creds_dict["private_key"] = "-----BEGIN PRIVATE KEY-----\n" + "\n".join(chunks) + "\n-----END PRIVATE KEY-----\n"

            scopes = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
            credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
            return gspread.authorize(credentials)
    except Exception:
        return None
    return None


def load_all_data():
    def_cars, def_rules, def_cons, def_reps, def_fbs, def_man = get_default_dataframes()

    # 1. 구글 시트 모드 켜져 있을 때 시도
    if ENABLE_GOOGLE_SHEETS:
        client = get_gspread_client()
        if client and "spreadsheet_url" in st.secrets:
            try:
                spreadsheet = client.open_by_url(st.secrets["spreadsheet_url"])
                d_cars = normalize_dataframe(pd.DataFrame(spreadsheet.worksheet("1_차량기본정보").get_all_records()), def_cars)
                d_rules = normalize_dataframe(pd.DataFrame(spreadsheet.worksheet("2_교체주기기준").get_all_records()), def_rules)
                d_cons = normalize_dataframe(pd.DataFrame(spreadsheet.worksheet("3_소모품이력").get_all_records()), def_cons)
                d_reps = normalize_dataframe(pd.DataFrame(spreadsheet.worksheet("4_정비수리이력").get_all_records()), def_reps)
                d_fbs = normalize_dataframe(pd.DataFrame(spreadsheet.worksheet("5_특이사항").get_all_records()), def_fbs)
                d_man = normalize_dataframe(pd.DataFrame(spreadsheet.worksheet("6_사고대처매뉴얼").get_all_records()), def_man)
                return (d_cars, d_rules, d_cons, d_reps, d_fbs, d_man), "🟢 구글 스프레드시트 실시간 연동중"
            except Exception:
                pass

    # 2. 로컬 엑셀 파일 로드 (없거나 구버전이면 자동 갱신)
    if os.path.exists(EXCEL_FILE):
        try:
            xls = pd.ExcelFile(EXCEL_FILE)
            d_cars = normalize_dataframe(pd.read_excel(xls, "1_차량기본정보"), def_cars)
            d_rules = normalize_dataframe(pd.read_excel(xls, "2_교체주기기준"), def_rules)
            d_cons = normalize_dataframe(pd.read_excel(xls, "3_소모품이력"), def_cons)
            d_reps = normalize_dataframe(pd.read_excel(xls, "4_정비수리이력"), def_reps)
            d_fbs = normalize_dataframe(pd.read_excel(xls, "5_특이사항"), def_fbs)
            d_man = normalize_dataframe(pd.read_excel(xls, "6_사고대처매뉴얼"), def_man)
            return (d_cars, d_rules, d_cons, d_reps, d_fbs, d_man), "💻 로컬 연동모드 (엑셀)"
        except Exception:
            pass

    # 3. 초기 엑셀 파일 생성
    save_all_data(def_cars, def_rules, def_cons, def_reps, def_fbs, def_man)
    return (def_cars, def_rules, def_cons, def_reps, def_fbs, def_man), "💻 로컬 연동모드 (엑셀)"


def save_all_data(df_cars, df_rules, df_cons, df_reps, df_fbs, df_man):
    # 로컬 엑셀 저장
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_cars.to_excel(writer, sheet_name="1_차량기본정보", index=False)
        df_rules.to_excel(writer, sheet_name="2_교체주기기준", index=False)
        df_cons.to_excel(writer, sheet_name="3_소모품이력", index=False)
        df_reps.to_excel(writer, sheet_name="4_정비수리이력", index=False)
        df_fbs.to_excel(writer, sheet_name="5_특이사항", index=False)
        df_man.to_excel(writer, sheet_name="6_사고대처매뉴얼", index=False)

    # 구글 시트 연동 시 동시 저장
    if ENABLE_GOOGLE_SHEETS:
        client = get_gspread_client()
        if client and "spreadsheet_url" in st.secrets:
            try:
                spreadsheet = client.open_by_url(st.secrets["spreadsheet_url"])
                sheet_map = {
                    "1_차량기본정보": df_cars,
                    "2_교체주기기준": df_rules,
                    "3_소모품이력": df_cons,
                    "4_정비수리이력": df_reps,
                    "5_특이사항": df_fbs,
                    "6_사고대처매뉴얼": df_man
                }
                for ws_name, df_target in sheet_map.items():
                    ws = spreadsheet.worksheet(ws_name)
                    ws.clear()
                    df_str = df_target.fillna("").astype(str)
                    ws.update([df_str.columns.values.tolist()] + df_str.values.tolist())
            except Exception as e:
                st.warning(f"구글 시트 동기화 중 오류: {e}")


# ---------------------------------------------------------
# 4. 유틸리티 계산 함수 (소모품 상태 및 D-Day 계산)
# ---------------------------------------------------------
def safe_int(val, default=0):
    try:
        if pd.isna(val) or str(val).strip() == "":
            return default
        return int(float(str(val).replace(",", "").strip()))
    except Exception:
        return default


def evaluate_consumable_status(current_km, last_km, cycle_km, last_date_str, cycle_months):
    driven_km = max(0, safe_int(current_km) - safe_int(last_km))
    cycle_km_val = max(1, safe_int(cycle_km, 10000))
    km_ratio = driven_km / cycle_km_val

    months_elapsed = 0.0
    try:
        last_dt = datetime.strptime(str(last_date_str).strip()[:10], "%Y-%m-%d").date()
        days_diff = (date.today() - last_dt).days
        months_elapsed = max(0.0, days_diff / 30.4)
    except Exception:
        months_elapsed = 0.0

    cycle_m_val = max(1, safe_int(cycle_months, 12))
    month_ratio = months_elapsed / cycle_m_val
    max_ratio = max(km_ratio, month_ratio)

    if max_ratio >= 1.0:
        badge = "🔴 교체필요"
    elif max_ratio >= 0.8:
        badge = "🟡 점검요망"
    else:
        badge = "🟢 양호"

    return badge, driven_km, round(months_elapsed, 1), min(1.0, max_ratio)


def calc_dday_badge(target_date_str):
    try:
        target_dt = datetime.strptime(str(target_date_str).strip()[:10], "%Y-%m-%d").date()
        diff = (target_dt - date.today()).days
        if diff < 0:
            return f"🔴 만료 ({abs(diff)}일 경과)", diff
        elif diff <= 30:
            return f"🟡 임박 (D-{diff})", diff
        else:
            return f"🟢 정상 (D-{diff})", diff
    except Exception:
        return "⚪ 미상", 999


# ---------------------------------------------------------
# 5. 데이터 불러오기 및 사이드바 구성
# ---------------------------------------------------------
(df_cars, df_rules, df_cons, df_reps, df_fbs, df_man), db_status_msg = load_all_data()

st.sidebar.title("🚘 차량 QR 관리 시스템")
st.sidebar.caption(db_status_msg)
if st.sidebar.button("🔄 데이터 새로고침", use_container_width=True):
    st.rerun()

st.sidebar.divider()

# URL 쿼리 파라미터 확인 (?car=본사-1)
query_params = st.query_params
qr_car_param = query_params.get("car", None)

view_mode = st.sidebar.radio(
    "🖥️ 화면 모드 선택",
    ["📱 모바일 QR 화면 (현장)", "📊 관리자 대시보드 (보고용)"],
    index=0
)

st.sidebar.divider()

# ---------------------------------------------------------
# 6. 화면 모드 1: 모바일 QR 화면 (현장)
# ---------------------------------------------------------
if view_mode == "📱 모바일 QR 화면 (현장)":
    st.sidebar.subheader("🔍 QR 스캔 차량 선택 (2단계)")

    all_branches = sorted(df_cars["소속본부"].dropna().unique().tolist())
    default_branch_idx = 0
    default_car_name = None

    if qr_car_param and qr_car_param in df_cars["차량명"].values:
        matched_branch = df_cars.loc[df_cars["차량명"] == qr_car_param, "소속본부"].iloc[0]
        if matched_branch in all_branches:
            default_branch_idx = all_branches.index(matched_branch)
            default_car_name = qr_car_param

    sel_branch = st.sidebar.selectbox("1단계: 소속 본부 선택", all_branches, index=default_branch_idx)
    branch_cars_df = df_cars[df_cars["소속본부"] == sel_branch]
    car_list = branch_cars_df["차량명"].tolist()

    default_car_idx = 0
    if default_car_name and default_car_name in car_list:
        default_car_idx = car_list.index(default_car_name)

    sel_car = st.sidebar.selectbox(
        "2단계: 차량 선택",
        car_list,
        index=default_car_idx,
        format_func=lambda x: f"{x} ({branch_cars_df.loc[branch_cars_df['차량명']==x, '차종'].iloc[0]})"
    )

    if sel_car:
        car_row = df_cars[df_cars["차량명"] == sel_car].iloc[0]
        current_km = safe_int(car_row.get("최근 주행거리", 0))
        insp_badge, _ = calc_dday_badge(car_row.get("정기검사 만료일", ""))
        ins_badge, _ = calc_dday_badge(car_row.get("보험만료일", ""))

        # 상단 네이비 차량 배너
        st.markdown(f"""
        <div class="car-header-banner">
            <div class="car-header-title">🚗 {car_row['차량명']} ({car_row['차종']})</div>
            <div class="car-header-sub">
                <b>차량번호:</b> {car_row['차량정보']} &nbsp;|&nbsp;
                <b>소속:</b> {car_row['소속본부']} ({car_row['용도구분']}) &nbsp;|&nbsp;
                <b>연료:</b> {car_row['연료']} &nbsp;|&nbsp;
                <b>담당:</b> {car_row['담당자_연락처']}
            </div>
        </div>
        """, unsafe_allow_html=True)

        # 차량 핵심 상태 요약 바 & 주행거리 빠른 갱신
        c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.2, 1.4])
        c1.metric("현재 누적 주행거리", f"{current_km:,} km")
        c2.metric("정기검사 만료일", str(car_row.get("정기검사 만료일", "-")), insp_badge)
        c3.metric("자동차 보험만료일", str(car_row.get("보험만료일", "-")), ins_badge)

        with c4:
            with st.popover("📍 주행거리 즉시 업데이트", use_container_width=True):
                new_km_val = st.number_input("계기판 현재 주행거리 (km)", min_value=0, value=current_km, step=100)
                if st.button("주행거리 저장", type="primary", use_container_width=True):
                    df_cars.loc[df_cars["차량명"] == sel_car, "최근 주행거리"] = int(new_km_val)
                    save_all_data(df_cars, df_rules, df_cons, df_reps, df_fbs, df_man)
                    st.toast("주행거리가 업데이트되었습니다!")
                    st.rerun()

        # 3대 메인 탭 구성
        tab_maint, tab_fb, tab_manual = st.tabs(["🛠️ 정비·소모품 이력", "💬 특이사항", "🚨 사고절차"])

        # [탭 1] 정비·소모품 이력
        with tab_maint:
            st.subheader("🔧 소모품 교체 주기 및 상태 점검")
            car_cons = df_cons[df_cons["차량명"] == sel_car]

            if car_cons.empty:
                st.info("등록된 소모품 관리 항목이 없습니다.")
            else:
                cols_cons = st.columns(2)
                for idx, (_, c_row) in enumerate(car_cons.iterrows()):
                    badge, driven_km, m_elapsed, ratio = evaluate_consumable_status(
                        current_km,
                        c_row["마지막 교체 주행거리"],
                        c_row["교체주기(Km)"],
                        c_row["마지막 교체일"],
                        c_row["교체주기(개월)"]
                    )
                    with cols_cons[idx % 2]:
                        st.markdown(f"""
                        <div class="status-card">
                            <b>{c_row['항목']}</b> &nbsp; <span style="float:right; font-weight:bold;">{badge}</span><br>
                            <small style="color:#475569;">
                                마지막 교체: {c_row['마지막 교체일']} ({safe_int(c_row['마지막 교체 주행거리']):,} km)<br>
                                경과: <b>+{driven_km:,} km</b> / 기준 {safe_int(c_row['교체주기(Km)']):,} km ({m_elapsed}개월 경과)
                            </small>
                        </div>
                        """, unsafe_allow_html=True)
                        st.progress(ratio)

            st.divider()
            st.subheader("📜 주요 정비 및 수리 이력")
            car_reps = df_reps[df_reps["차량명"] == sel_car].sort_values(by="일자", ascending=False)

            if car_reps.empty:
                st.caption("기록된 정비·수리 이력이 없습니다.")
            else:
                for _, r_row in car_reps.iterrows():
                    st.markdown(
                        f"- **[{r_row['정비구분']}] {r_row['내용']}** "
                        f"(일자: `{r_row['일자']}` | 주행거리: `{safe_int(r_row['주행거리(Km)']):,} km` | 정비소: `{r_row['정비소']}`)"
                    )

            with st.expander("➕ 현장 정비·소모품 교체 내역 신규 등록"):
                with st.form("add_repair_form", clear_on_submit=True):
                    f_col1, f_col2 = st.columns(2)
                    r_date = f_col1.date_input("정비 일자", value=date.today())
                    r_km = f_col2.number_input("정비 시점 주행거리 (Km)", min_value=0, value=current_km, step=100)
                    r_type = f_col1.selectbox("정비 구분", ["소모품", "일반정비", "사고수리", "정기점검"])
                    r_item = f_col2.selectbox("연동할 소모품 항목 (소모품 교체 시)", ["선택 안함"] + df_rules["항목"].tolist())
                    r_desc = st.text_input("정비 상세 내용 (예: 엔진오일 및 에어컨필터 교체)")
                    r_shop = st.text_input("정비소명 (예: 블루핸즈 분당점)")

                    if st.form_submit_button("정비 이력 저장", type="primary"):
                        if not r_desc.strip():
                            st.warning("정비 상세 내용을 입력해 주세요.")
                        else:
                            new_rep = pd.DataFrame([{
                                "차량명": sel_car,
                                "일자": r_date.strftime("%Y-%m-%d"),
                                "주행거리(Km)": int(r_km),
                                "정비구분": r_type,
                                "내용": r_desc.strip(),
                                "정비소": r_shop.strip() if r_shop.strip() else "-"
                            }])
                            df_reps = pd.concat([df_reps, new_rep], ignore_index=True)

                            # 주행거리 최신화
                            if r_km > current_km:
                                df_cars.loc[df_cars["차량명"] == sel_car, "최근 주행거리"] = int(r_km)

                            # 소모품 이력 자동 연동 업데이트
                            if r_item != "선택 안함":
                                mask = (df_cons["차량명"] == sel_car) & (df_cons["항목"] == r_item)
                                if mask.any():
                                    df_cons.loc[mask, "마지막 교체일"] = r_date.strftime("%Y-%m-%d")
                                    df_cons.loc[mask, "마지막 교체 주행거리"] = int(r_km)
                                else:
                                    rule_row = df_rules[df_rules["항목"] == r_item].iloc[0]
                                    new_con = pd.DataFrame([{
                                        "차량명": sel_car,
                                        "항목": r_item,
                                        "교체주기(Km)": int(rule_row["권장주기_km"]),
                                        "교체주기(개월)": int(rule_row["권장주기_개월"]),
                                        "마지막 교체일": r_date.strftime("%Y-%m-%d"),
                                        "마지막 교체 주행거리": int(r_km)
                                    }])
                                    df_cons = pd.concat([df_cons, new_con], ignore_index=True)

                            save_all_data(df_cars, df_rules, df_cons, df_reps, df_fbs, df_man)
                            st.success("정비 이력이 정상 등록되었습니다!")
                            st.rerun()

        # [탭 2] 특이사항
        with tab_fb:
            st.subheader("💬 차량 특이사항 및 이상 징후 제보")
            car_fbs = df_fbs[df_fbs["차량명"] == sel_car].sort_values(by="날짜", ascending=False)

            if car_fbs.empty:
                st.info("접수된 차량 특이사항이 없습니다.")
            else:
                st.dataframe(car_fbs, use_container_width=True, hide_index=True)

            st.markdown("#### ✍️ 운행 중 특이사항 신규 등록")
            with st.form("add_feedback_form", clear_on_submit=True):
                fb_c1, fb_c2 = st.columns(2)
                fb_writer = fb_c1.text_input("작성자 (성명/직급)")
                fb_cat = fb_c2.selectbox("분류", ["소음/진동", "경고등 점등", "외관손상", "타이어/제동", "실내청결/비품", "기타"])
                fb_content = st.text_area("상세 내용 (증상이나 요청사항을 적어주세요)")

                if st.form_submit_button("특이사항 등록", type="primary"):
                    if not fb_writer.strip() or not fb_content.strip():
                        st.warning("작성자와 상세 내용을 모두 입력해 주세요.")
                    else:
                        new_fb = pd.DataFrame([{
                            "차량명": sel_car,
                            "날짜": date.today().strftime("%Y-%m-%d"),
                            "작성자": fb_writer.strip(),
                            "확인일자": "-",
                            "분류": fb_cat,
                            "내용": fb_content.strip(),
                            "처리상태": "🔴 확인필요"
                        }])
                        df_fbs = pd.concat([df_fbs, new_fb], ignore_index=True)
                        save_all_data(df_cars, df_rules, df_cons, df_reps, df_fbs, df_man)
                        st.success("특이사항이 담당 부서로 접수되었습니다!")
                        st.rerun()

        # [탭 3] 사고절차
        with tab_manual:
            st.subheader("🚨 사고 및 긴급상황 발생 시 행동 매뉴얼")
            st.warning(f"📞 **긴급 연락처 안내** | 차량 담당 부서: **{car_row['담당자_연락처']}**")

            for _, m_row in df_man.sort_values(by="순서").iterrows():
                st.markdown(f"""
                <div class="status-card">
                    <h4 style="margin:0 0 6px 0; color:#1e3c72;">{m_row['단계명']}</h4>
                    <p style="margin:0; font-size:15px;">{m_row['행동요령']}</p>
                </div>
                """, unsafe_allow_html=True)


# ---------------------------------------------------------
# 7. 화면 모드 2: 관리자 대시보드 (보고용)
# ---------------------------------------------------------
else:
    branch_list = ["전체 보기"] + sorted(df_cars["소속본부"].dropna().unique().tolist())
    selected_branch = st.sidebar.selectbox("조회 대상 본부", branch_list)

    st.title("📊 전사 차량 QR 이력·상태 통합 관제 대시보드")

    filtered_cars = df_cars if selected_branch == "전체 보기" else df_cars[df_cars["소속본부"] == selected_branch]
    target_car_names = filtered_cars["차량명"].tolist()

    # 소모품 점검/교체 필요 건수 집계
    alert_cons_list = []
    for _, c_row in df_cons[df_cons["차량명"].isin(target_car_names)].iterrows():
        c_km = safe_int(df_cars.loc[df_cars["차량명"] == c_row["차량명"], "최근 주행거리"].iloc[0])
        badge, driven_km, m_elapsed, _ = evaluate_consumable_status(
            c_km, c_row["마지막 교체 주행거리"], c_row["교체주기(Km)"], c_row["마지막 교체일"], c_row["교체주기(개월)"]
        )
        if "🔴" in badge or "🟡" in badge:
            alert_cons_list.append({
                "차량명": c_row["차량명"],
                "항목": c_row["항목"],
                "상태": badge,
                "경과주행(Km)": driven_km,
                "경과개월": m_elapsed
            })

    # 미조치 특이사항 건수 집계
    filtered_fbs = df_fbs[df_fbs["차량명"].isin(target_car_names)]
    pending_fbs_cnt = len(filtered_fbs[filtered_fbs["처리상태"].astype(str).str.contains("확인필요|조치중")])

    # 검사/보험 만료 임박 건수 집계
    exp_alert_cnt = 0
    for _, car_r in filtered_cars.iterrows():
        _, d1 = calc_dday_badge(car_r.get("정기검사 만료일", ""))
        _, d2 = calc_dday_badge(car_r.get("보험만료일", ""))
        if d1 <= 30 or d2 <= 30:
            exp_alert_cnt += 1

    # 상단 KPI 요약 카드 4개
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("🚘 관리 대상 차량", f"{len(filtered_cars)} 대")
    k2.metric("🔧 소모품 점검·교체 도래", f"{len(alert_cons_list)} 건")
    k3.metric("💬 미조치 현장 특이사항", f"{pending_fbs_cnt} 건")
    k4.metric("📅 검사·보험 만료 임박(30일내)", f"{exp_alert_cnt} 대")

    st.divider()

    adm_tab1, adm_tab2, adm_tab3, adm_tab4 = st.tabs([
        "📋 차량 현황 및 소모품 관제",
        "💬 특이사항 접수·조치 관리",
        "✏️ 전체 데이터 통합 편집 (DB 관리)",
        "📱 차량별 QR 코드 생성·인쇄"
    ])

    # [관리자 탭 1] 차량 현황 및 소모품 관제
    with adm_tab1:
        st.subheader("🚘 전사 차량 기본 현황")
        st.dataframe(filtered_cars, use_container_width=True, hide_index=True)

        col_a1, col_a2 = st.columns(2)
        with col_a1:
            st.subheader("⚠️ 소모품 교체·점검 도래 현황")
            if alert_cons_list:
                st.dataframe(pd.DataFrame(alert_cons_list), use_container_width=True, hide_index=True)
            else:
                st.success("모든 차량의 소모품 상태가 양호합니다.")

        with col_a2:
            st.subheader("🛠️ 최근 정비·수리 내역")
            recent_reps = df_reps[df_reps["차량명"].isin(target_car_names)].sort_values(by="일자", ascending=False)
            st.dataframe(recent_reps, use_container_width=True, hide_index=True)

    # [관리자 탭 2] 특이사항 접수·조치 관리
    with adm_tab2:
        st.subheader("💬 현장 접수 특이사항 실시간 조치 관리")
        st.caption("표에서 '처리상태'와 '확인일자'를 직접 수정한 뒤 아래 저장 버튼을 누르면 즉시 반영됩니다.")
        edited_fbs = st.data_editor(
            df_fbs,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "처리상태": st.column_config.SelectboxColumn(
                    "처리상태",
                    options=["🔴 확인필요", "🟡 조치중", "🟢 조치완료"],
                    required=True
                )
            },
            key="editor_fbs_status"
        )
        if st.button("💾 특이사항 조치 내역 저장", type="primary"):
            df_fbs = edited_fbs
            save_all_data(df_cars, df_rules, df_cons, df_reps, df_fbs, df_man)
            st.success("특이사항 상태가 저장되었습니다!")
            st.rerun()

    # [관리자 탭 3] 전체 데이터 통합 편집기
    with adm_tab3:
        st.subheader("✏️ 6대 시트 마스터 데이터 직접 편집")
        sheet_choice = st.selectbox(
            "편집할 시트 선택",
            ["1_차량기본정보", "2_교체주기기준", "3_소모품이력", "4_정비수리이력", "5_특이사항", "6_사고대처매뉴얼"]
        )

        if sheet_choice == "1_차량기본정보":
            ed_df = st.data_editor(df_cars, num_rows="dynamic", use_container_width=True, hide_index=True)
        elif sheet_choice == "2_교체주기기준":
            ed_df = st.data_editor(df_rules, num_rows="dynamic", use_container_width=True, hide_index=True)
        elif sheet_choice == "3_소모품이력":
            ed_df = st.data_editor(df_cons, num_rows="dynamic", use_container_width=True, hide_index=True)
        elif sheet_choice == "4_정비수리이력":
            ed_df = st.data_editor(df_reps, num_rows="dynamic", use_container_width=True, hide_index=True)
        elif sheet_choice == "5_특이사항":
            ed_df = st.data_editor(df_fbs, num_rows="dynamic", use_container_width=True, hide_index=True)
        else:
            ed_df = st.data_editor(df_man, num_rows="dynamic", use_container_width=True, hide_index=True)

        if st.button("💾 선택한 시트 변경사항 저장", type="primary"):
            if sheet_choice == "1_차량기본정보":
                df_cars = ed_df
            elif sheet_choice == "2_교체주기기준":
                df_rules = ed_df
            elif sheet_choice == "3_소모품이력":
                df_cons = ed_df
            elif sheet_choice == "4_정비수리이력":
                df_reps = ed_df
            elif sheet_choice == "5_특이사항":
                df_fbs = ed_df
            else:
                df_man = ed_df

            save_all_data(df_cars, df_rules, df_cons, df_reps, df_fbs, df_man)
            st.success(f"'{sheet_choice}' 데이터가 저장되었습니다!")
            st.rerun()

    # [관리자 탭 4] 차량별 QR 코드 생성
    with adm_tab4:
        st.subheader("📱 차량 부착용 다이렉트 QR 코드 생성기")
        base_url = st.text_input("배포된 스트림릿 앱 주소 입력", value="https://k-cars.streamlit.app")
        qr_cols = st.columns(4)
        for idx, (_, c_row) in enumerate(filtered_cars.iterrows()):
            c_name = c_row["차량명"]
            encoded_car = urllib.parse.quote(str(c_name))
            target_link = f"{base_url.rstrip('/')}/?car={encoded_car}"
            qr_img_url = f"https://api.qrserver.com/v1/create-qr-code/?size=200x200&data={urllib.parse.quote(target_link)}"
            with qr_cols[idx % 4]:
                st.markdown(f"**🚗 {c_name}** ({c_row['차종']})")
                st.image(qr_img_url, width=160)
                st.caption(f"차량번호: {c_row['차량정보']}")