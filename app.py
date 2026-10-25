import streamlit as st
import pandas as pd
from datetime import datetime
import urllib.parse
import os

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

st.set_page_config(
    page_title="차량 QR 이력·상태 관리 시스템",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded"
)

EXCEL_FILE = "car_data.xlsx"

def get_gspread_client():
    if not GSPREAD_AVAILABLE:
        return None
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            if "private_key" in creds_dict:
                raw_key = str(creds_dict["private_key"])
                
                # PEM 키 정제 및 패딩 검사
                cleaned = raw_key.replace("-----BEGIN PRIVATE KEY-----", "")
                cleaned = cleaned.replace("-----END PRIVATE KEY-----", "")
                cleaned = cleaned.replace("\\n", "\n").replace("\r", "")
                pure_chars = "".join(cleaned.split())
                
                # 만약 키가 잘렸거나 패딩이 깨졌는지 여부 확인을 위해 디버그 메시지 출력 가능
                chunks = [pure_chars[i:i+64] for i in range(0, len(pure_chars), 64)]
                formatted_key = "-----BEGIN PRIVATE KEY-----\n" + "\n".join(chunks) + "\n-----END PRIVATE KEY-----\n"
                
                creds_dict["private_key"] = formatted_key
                
            scopes = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
            credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
            client = gspread.authorize(credentials)
            return client
    except Exception as e:
        st.sidebar.error(f"구글 인증 실패 디버그: {e}")
    return None

def load_data_from_google_sheets():
    client = get_gspread_client()
    if not client:
        return None, None
    try:
        if "spreadsheet_url" in st.secrets:
            sheet_url = st.secrets["spreadsheet_url"]
            spreadsheet = client.open_by_url(sheet_url)
            
            d_cars = pd.DataFrame(spreadsheet.worksheet("1_차량기본정보").get_all_records())
            d_rules = pd.DataFrame(spreadsheet.worksheet("2_교체주기기준").get_all_records())
            d_cons = pd.DataFrame(spreadsheet.worksheet("3_소모품이력").get_all_records())
            d_reps = pd.DataFrame(spreadsheet.worksheet("4_정비수리이력").get_all_records())
            d_fbs = pd.DataFrame(spreadsheet.worksheet("5_특이사항").get_all_records())
            d_man = pd.DataFrame(spreadsheet.worksheet("6_사고대처매뉴얼").get_all_records())
            
            return (d_cars, d_rules, d_cons, d_reps, d_fbs, d_man), "🟢 구글 스프레드시트 실시간 연동중"
    except Exception as e:
        return None, f"⚠️ 구글 시트 연동 오류: {e}"
    return None, None

def get_default_dataframes():
    df_cars = pd.DataFrame([
        {"차량명": "본사-1", "차종": "아이오닉 6 (전기차)", "차량정보": "88하 1024", "소속본부": "본사 (지원처)", "용도구분": "일반 업무용", "연료": "전기(EV)", "담당자_연락처": "본사 지원처", "최근 주행거리": 24150, "정기검사 만료일": "2027-03-01", "보험만료일": "2027-02-15"},
    ])
    df_rules = pd.DataFrame([{"항목": "엔진오일", "권장주기_km": 10000, "권장주기_개월": 12}])
    df_cons = pd.DataFrame([{"차량명": "본사-1", "항목": "엔진오일", "교체주기(Km)": 10000, "교체주기(개월)": 12, "마지막 교체일": "2026-01-15", "마지막 교체 주행거리": 14000}])
    df_reps = pd.DataFrame([{"차량명": "본사-1", "일자": "2026-07-20", "주행거리(Km)": 14000, "정비구분": "소모품", "내용": "엔진오일 교체", "정비소": "블루핸즈"}])
    df_fbs = pd.DataFrame([{"차량명": "본사-1", "날짜": "2026-10-22", "작성자": "김주임", "확인일자": "2026-10-22", "분류": "기타", "내용": "점검 완료", "처리상태": "🟢 조치완료"}])
    df_man = pd.DataFrame([{"순서": 1, "단계명": "1단계: 안전 확보", "행동요령": "비상등 켜고 안전지대로 대피합니다."}])
    return df_cars, df_rules, df_cons, df_reps, df_fbs, df_man

gs_data, db_status_msg = load_data_from_google_sheets()
if gs_data is not None:
    df_cars, df_rules, df_cons, df_reps, df_fbs, df_man = gs_data
else:
    if not os.path.exists(EXCEL_FILE):
        d_cars, d_rules, d_cons, d_reps, d_fbs, d_man = get_default_dataframes()
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            for s_name, df in {"1_차량기본정보": d_cars, "2_교체주기기준": d_rules, "3_소모품이력": d_cons, "4_정비수리이력": d_reps, "5_특이사항": d_fbs, "6_사고대처매뉴얼": d_man}.items():
                df.to_excel(writer, sheet_name=s_name, index=False)
    xls = pd.ExcelFile(EXCEL_FILE)
    df_cars = pd.read_excel(xls, "1_차량기본정보")
    df_rules = pd.read_excel(xls, "2_교체주기기준")
    df_cons = pd.read_excel(xls, "3_소모품이력")
    df_reps = pd.read_excel(xls, "4_정비수리이력")
    df_fbs = pd.read_excel(xls, "5_특이사항")
    df_man = pd.read_excel(xls, "6_사고대처매뉴얼")
    db_status_msg = "💻 로컬 연동모드 (엑셀)"

st.sidebar.title("🚘 차량 QR 관리 시스템")
st.sidebar.caption(db_status_msg)
if st.sidebar.button("🔄 데이터 새로고침", use_container_width=True):
    st.rerun()

st.sidebar.divider()
branch_list = ["전체 보기"] + sorted(df_cars["소속본부"].dropna().unique().tolist())
selected_branch = st.sidebar.selectbox("조회 대상 본부", branch_list)

st.title("📊 전사 차량 QR 이력·상태 통합 관제 대시보드")
st.dataframe(df_cars, use_container_width=True, hide_index=True)