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
    .popup-alert-box {
        background-color: #fff5f5;
        border-left: 5px solid #e53e3e;
        padding: 12px 15px;
        border-radius: 6px;
        margin-bottom: 10px;
    }
    .popup-warn-box {
        background-color: #fefcbf;
        border-left: 5px solid #d69e2e;
        padding: 12px 15px;
        border-radius: 6px;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

EXCEL_FILE = "car_data.xlsx"

# 구글 시트 실시간 연동 활성화 및 기본 시트 주소 내장
ENABLE_GOOGLE_SHEETS = True
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HDP5OwqbtV4xRerESbj1_inMNLtiFA8Q7QTC6ThB7Kk/edit"


# ---------------------------------------------------------
# 2. 초기 기본 데이터 정의
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
            "정기검사 만료일": "2027-03-01"
        },
        {
            "차량명": "본사-2",
            "차종": "쏘나타 DN8",
            "차량정보": "123가 4567",
            "소속본부": "본사 (지원처)",
            "용도구분": "일반 업무용",
            "연료": "가솔린",
            "담당자_연락처": "본사 지원처 (내선 0000)",
            "최근 주행거리": 49500,
            "정기검사 만료일": "2026-12-15"
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
            "정기검사 만료일": "2026-11-15"
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
            "정기검사 만료일": "2027-04-10"
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
        {"차량명": "본사-1", "항목": "에어컨필터", "교체주기(Km)": 15000, "교체주기(개월)": 6, "마지막 교체일": "2026-03-15", "마지막 교체 주행거리": 10000},
        {"차량명": "본사-1", "항목": "타이어 위치교환/교체", "교체주기(Km)": 20000, "교체주기(개월)": 24, "마지막 교체일": "2025-11-10", "마지막 교체 주행거리": 10000},
        {"차량명": "본사-1", "항목": "브레이크패드", "교체주기(Km)": 40000, "교체주기(개월)": 24, "마지막 교체일": "2025-05-01", "마지막 교체 주행거리": 5000},
        {"차량명": "본사-1", "항목": "와이퍼 블레이드", "교체주기(Km)": 10000, "교체주기(개월)": 12, "마지막 교체일": "2026-07-20", "마지막 교체 주행거리": 19500},
        {"차량명": "본사-1", "항목": "12V 보조배터리", "교체주기(Km)": 50000, "교체주기(개월)": 36, "마지막 교체일": "2025-01-10", "마지막 교체 주행거리": 0},
        {"차량명": "본사-2", "항목": "엔진오일", "교체주기(Km)": 10000, "교체주기(개월)": 12, "마지막 교체일": "2026-01-10", "마지막 교체 주행거리": 40000},
        {"차량명": "본사-2", "항목": "에어컨필터", "교체주기(Km)": 15000, "교체주기(개월)": 6, "마지막 교체일": "2026-03-10", "마지막 교체 주행거리": 40000},
        {"차량명": "북부-1", "항목": "엔진오일", "교체주기(Km)": 10000, "교체주기(개월)": 12, "마지막 교체일": "2025-12-05", "마지막 교체 주행거리": 50500},
        {"차량명": "영남-1", "항목": "엔진오일", "교체주기(Km)": 10000, "교체주기(개월)": 12, "마지막 교체일": "2026-05-20", "마지막 교체 주행거리": 82000}
    ])

    df_reps = pd.DataFrame([
        {"차량명": "본사-1", "일자": "2026-07-20", "주행거리(Km)": 19500, "정비구분": "소모품", "내용": "와이퍼 블레이드 교체", "정비소": "블루핸즈 분당점"},
        {"차량명": "본사-2", "일자": "2026-03-10", "주행거리(Km)": 40000, "정비구분": "소모품", "내용": "에어컨필터 교체", "정비소": "블루핸즈 수지점"},
        {"차량명": "북부-1", "일자": "2026-08-12", "주행거리(Km)": 58000, "정비구분": "일반정비", "내용": "타이어 펑크 패치 수리", "정비소": "오토큐 의정부점"}
    ])

    df_fbs = pd.DataFrame([
        {"차량명": "본사-1", "날짜": "2026-10-25", "작성자": "박대리 (지원처)", "확인일자": "-", "분류": "타이어/제동", "내용": "운전석 앞바퀴 공기압 경고등 점등 확인 필요", "처리상태": "조치필요"},
        {"차량명": "북부-1", "날짜": "2026-10-24", "작성자": "이대리 (북부운영)", "확인일자": "-", "분류": "소음/진동", "내용": "방지턱 넘을 때 하부 소음 약간 발생", "처리상태": "조치필요"}
    ])

    df_man = pd.DataFrame([
        {"순서": 1, "단계명": "1단계: 비상등 점등 및 안전 확보", "행동요령": "사고 즉시 비상등을 켜고, 차량 이동이 가능하면 갓길 등 안전지대로 대피합니다."},
        {"순서": 2, "단계명": "2단계: 부상자 구호 및 119/112 신고", "행동요령": "동승자 및 상대방 부상 여부를 최우선 확인하고 필요시 119 및 경찰(112)에 신고합니다."},
        {"순서": 3, "단계명": "3단계: 현장 사진 촬영 및 증거 확보", "행동요령": "원거리(차선 포함), 근거리(파손 부위), 상대 차량 번호판 및 블랙박스 영상을 확보합니다."},
        {"순서": 4, "단계명": "4단계: 보험사 접수 및 담당 부서 보고", "행동요령": "공용차량 전담 보험사 콜센터에 사고를 접수하고 소속 본부 차량 담당자에게 즉시 보고합니다."}
    ])

    return df_cars, df_rules, df_cons, df_reps, df_fbs, df_man


# ---------------------------------------------------------
# 3. 데이터 무결성 보장 및 로드/저장 함수
# ---------------------------------------------------------
def normalize_dataframe(df, default_df, drop_legacy_cols=True):
    if df is None or df.empty:
        return default_df.copy()
    df_clean = df.copy()
    df_clean.columns = [str(c).strip() for c in df_clean.columns]
    if drop_legacy_cols:
        for legacy_col in ["주차위치", "주차 위치", "키위치", "키 위치", "보관장소", "보험만료일", "자동차 보험만료일"]:
            if legacy_col in df_clean.columns:
                df_clean = df_clean.drop(columns=[legacy_col])
    for col in default_df.columns:
        if col not in df_clean.columns:
            df_clean[col] = default_df[col].iloc[0] if len(default_df) > 0 else ""
    return df_clean[default_df.columns]


def get_sheet_url():
    if "spreadsheet_url" in st.secrets:
        return str(st.secrets["spreadsheet_url"]).strip()
    if "gcp_service_account" in st.secrets and "spreadsheet_url" in st.secrets["gcp_service_account"]:
        return str(st.secrets["gcp_service_account"]["spreadsheet_url"]).strip()
    return DEFAULT_SHEET_URL


def get_gspread_client():
    if not GSPREAD_AVAILABLE or not ENABLE_GOOGLE_SHEETS:
        return None, "연동 비활성화"
    try:
        # 백슬래시(\n) 이스케이프 문제를 원천 차단하기 위해 실제 줄바꿈 삼중 따옴표 사용
        raw_private_key = """-----BEGIN PRIVATE KEY-----
MIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDHyIeEndKxlZTv
maHnDReOhPyH3QNU/C46aVbltQeNJOTAe+DV8ak99ooszi78afvnTP0JLK6gCMU9
4WgalA8VwvcERYIrJT/l6MIJJ5LoOQToUcxRdWmtaZrXIz7NIQpTMih1Or7adiaD
xJ7i1BZGmVXr1VqCpA7tED5SIsMcFnseNCXUJwHStgfPbh3ssxONDjbiG18ciB4K
p7+0ZqTr7xEPpri5ETcc1SPC48LnUqlkaPGY47rKyr9LdJ/9RW1SHXXTful1fh+q
aEnPDoRFwXglgDK/vzMr90IyU6LW6WIWDcX+GMdCg3RoPXR+6xNRpqVsqY1/d1Ls
nd/fiBN1zAgMBAAECggEAClGKA6TNy0HSGhwqleDeVpqHxdH5y9usTq1bCII0qMA/
EFyq6tOs6hSY1X60Q+NuWC6hxp6kny2b5euiAxK/CgtdKIX8wVorK5hLwp8Pvnz6
iwU8wTEy0lBMrwjNp1jUoXdngxhMwzSsVsyYEuY2shlyOi3K1End0NJ3uuK7afDE
UEWouhwJbhKiT9+JXgVM5F7TQtSiHH5boq3RKqhqlroVAxRhFx/SIe2btMmkkSuv
xwylb9oTQPQcRliPaXD3LTcRhLlV6kasWXCWzxlnggTxs0a1wI6w8qK2H9u56/BN
nuOgoTCWK6+Od9qxamQh8xsb5A0Ri+TbyFPVSQDiIAQKBgQD1SRlQ42DE7vtZw/PY
/quK6Usjpf6xtJhPnMlp41eC2psJsL4HS70zT/G4980nGo9D1yZCT90ylgRsqz6V
H/YglfJRVg+MfFrCHvImiQtFVPhkofQRnjubNYvGprBOY5YpAjzg7Zgh0pcad7vo
nuX//CEbUopKAyjpUdHqm4OYpwKBgQDQgppJIEeVZvwhN3BLB8+VFbMJKPXZjlrj
mC3GhSy4LtLfP8/oOI2q7bb1GXlDLwego5TVMnljveX+m61nwVCY+TV08VHrfW4G
JCniOK+E7wl3mRifgXJndXiIy+bCH3WVsRXptrdfR0tWVnLLd7ghxreSeG+7mYoP
9tyh9ysiVQKBgQCNM3J2ZR6M38IkD7CtZxWbVudhcypmZD8V/XzZDv2On4eGkY8j
Jg+ZnVjakRQxgZGdjRpCLDV67G9Jm17ge9C1O3yOJAu5ZNDeVPAzDMtHMz32uGCH
T96iX7XCOqzdw2hq3RXsqfnc5p0DqI+svuRKc3sffmmLZ7xCbE56TuH3mwKBgCZl
Jkv46bWVNTBYqRDn602kpUc2h678C7iFOUipFUcFOEdr1LkKUOv8wgf4HOBGNL1b
4fgwRzIasdNkj1XfvxusKC/UkCUw/6WZ2yyRC6EK809WnbYhfBOZCg0GnwN4D4r1
BPf4cWPWGQVY/eKBtSJjGUKbcte0ujueVZptjT/RAoGAY+0yfYnuZPJGgt6AZTvA
vs1aVrWsy53lnGjZg+PXXDEPUY3iOq4SxfUjds4Pfmys4wv9yC+mmnGiwNlYcfaF
1wWrlU/obxkhp7nut8Net9rZncTjYjLwF56D+swplFgSMsg+J/eilrQBBAhMNG7P
Uvi+2j0lMHDGqGaK3sooxRc=
-----END PRIVATE KEY-----"""

        creds_info = {
            "type": "service_account",
            "project_id": "k-cars-512714",
            "private_key_id": "4f6e7e997b8b63d6b667bf5725b5aa8f6ae95f2a",
            "private_key": raw_private_key,
            "client_email": "k-cars@k-cars-512714.iam.gserviceaccount.com",
            "client_id": "101302661809874002306",
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/k-cars%40k-cars-512714.iam.gserviceaccount.com",
            "universe_domain": "googleapis.com"
        }
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        credentials = Credentials.from_service_account_info(creds_info, scopes=scopes)
        return gspread.authorize(credentials), None
    except Exception as e:
        return None, f"구글 인증 오류: {e}"


def read_worksheet_safe(spreadsheet, ws_name, default_df):
    try:
        ws = spreadsheet.worksheet(ws_name)
    except Exception:
        ws = spreadsheet.add_worksheet(title=ws_name, rows="100", cols="20")
        df_str = default_df.fillna("").astype(str)
        ws.update([df_str.columns.values.tolist()] + df_str.values.tolist())
        return default_df.copy()

    vals = ws.get_all_values()
    if not vals or len(vals) < 2:
        return default_df.copy()
    headers = [str(h).strip() for h in vals[0]]
    rows = [r for r in vals[1:] if any(str(cell).strip() != "" for cell in r)]
    if not rows:
        return pd.DataFrame(columns=default_df.columns)
    df_raw = pd.DataFrame(rows, columns=headers)
    return normalize_dataframe(df_raw, default_df)


def load_all_data():
    def_cars, def_rules, def_cons, def_reps, def_fbs, def_man = get_default_dataframes()
    gs_err_msg = None

    if ENABLE_GOOGLE_SHEETS:
        client, auth_err = get_gspread_client()
        if client:
            try:
                sheet_url = get_sheet_url()
                spreadsheet = client.open_by_url(sheet_url)
                d_cars = read_worksheet_safe(spreadsheet, "1_차량기본정보", def_cars)
                d_rules = read_worksheet_safe(spreadsheet, "2_교체주기기준", def_rules)
                d_cons = read_worksheet_safe(spreadsheet, "3_소모품이력", def_cons)
                d_reps = read_worksheet_safe(spreadsheet, "4_정비수리이력", def_reps)
                d_fbs = read_worksheet_safe(spreadsheet, "5_특이사항", def_fbs)
                d_man = read_worksheet_safe(spreadsheet, "6_사고대처매뉴얼", def_man)
                return (d_cars, d_rules, d_cons, d_reps, d_fbs, d_man), "🟢 구글 스프레드시트 실시간 연동중", None
            except Exception as e:
                gs_err_msg = f"시트 접근 오류(공유권한/URL 확인): {e}"
        else:
            gs_err_msg = auth_err

    if os.path.exists(EXCEL_FILE):
        try:
            xls = pd.ExcelFile(EXCEL_FILE)
            d_cars = normalize_dataframe(pd.read_excel(xls, "1_차량기본정보"), def_cars)
            d_rules = normalize_dataframe(pd.read_excel(xls, "2_교체주기기준"), def_rules)
            d_cons = normalize_dataframe(pd.read_excel(xls, "3_소모품이력"), def_cons)
            d_reps = normalize_dataframe(pd.read_excel(xls, "4_정비수리이력"), def_reps)
            d_fbs = normalize_dataframe(pd.read_excel(xls, "5_특이사항"), def_fbs)
            d_man = normalize_dataframe(pd.read_excel(xls, "6_사고대처매뉴얼"), def_man)
            return (d_cars, d_rules, d_cons, d_reps, d_fbs, d_man), "💻 로컬 연동모드 (엑셀)", gs_err_msg
        except Exception:
            pass

    save_all_data(def_cars, def_rules, def_cons, def_reps, def_fbs, def_man)
    return (def_cars, def_rules, def_cons, def_reps, def_fbs, def_man), "💻 로컬 연동모드 (엑셀)", gs_err_msg


def save_all_data(df_cars, df_rules, df_cons, df_reps, df_fbs, df_man):
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_cars.to_excel(writer, sheet_name="1_차량기본정보", index=False)
        df_rules.to_excel(writer, sheet_name="2_교체주기기준", index=False)
        df_cons.to_excel(writer, sheet_name="3_소모품이력", index=False)
        df_reps.to_excel(writer, sheet_name="4_정비수리이력", index=False)
        df_fbs.to_excel(writer, sheet_name="5_특이사항", index=False)
        df_man.to_excel(writer, sheet_name="6_사고대처매뉴얼", index=False)

    if ENABLE_GOOGLE_SHEETS:
        client, _ = get_gspread_client()
        if client:
            try:
                spreadsheet = client.open_by_url(get_sheet_url())
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
# 4. 유틸리티 계산 및 팝업(Dialog) 함수
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

    return badge, driven_km, round(months_elapsed), min(1.0, max_ratio)


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


@st.dialog("🚗 차량 상태정보 알림")
def show_car_alert_dialog(car_name, cons_alerts, fb_alerts, insp_alert):
    st.markdown(f"**[{car_name}]** 운행 전 반드시 아래 사항을 확인해 주세요.")

    if insp_alert:
        st.markdown(f"""
        <div class="popup-alert-box">
            <b>📅 정기검사 일정 알림</b><br>{insp_alert}
        </div>
        """, unsafe_allow_html=True)

    if cons_alerts:
        st.markdown("#### 🔧 소모품 교체·점검 도래 항목")
        for item in cons_alerts:
            box_cls = "popup-alert-box" if "🔴" in item["badge"] else "popup-warn-box"
            st.markdown(f"""
            <div class="{box_cls}">
                <b>{item['항목']} ({item['badge']})</b><br>
                <small>마지막 교체 후 <b>+{item['driven_km']:,} km</b> 주행 / <b>{item['m_elapsed']}개월</b> 경과 (기준: {item['cycle_km']:,} km / {item['cycle_m']}개월)</small>
            </div>
            """, unsafe_allow_html=True)

    if fb_alerts:
        st.markdown("#### 💬 미조치 차량 특이사항")
        for fb in fb_alerts:
            st.markdown(f"""
            <div class="popup-alert-box">
                <b>[{fb['분류']}] {fb['처리상태']}</b> (제보일: {fb['날짜']} / 작성자: {fb['작성자']})<br>
                <span>{fb['내용']}</span>
            </div>
            """, unsafe_allow_html=True)

    if st.button("확인완료", type="primary", use_container_width=True):
        st.rerun()


@st.dialog("전체 차량 상태정보 알림")
def show_admin_alert_dialog(branch_name, alert_cons_list, pending_fbs_df, exp_cars_list):
    if alert_cons_list:
        st.markdown(f"#### 🔧 소모품 교체·점검 도래 ({len(alert_cons_list)}건)")
        st.dataframe(pd.DataFrame(alert_cons_list), use_container_width=True, hide_index=True)

    if not pending_fbs_df.empty:
        st.markdown(f"#### 💬 미조치 현장 특이사항 ({len(pending_fbs_df)}건)")
        st.dataframe(
            pending_fbs_df[["차량명", "날짜", "작성자", "분류", "내용", "처리상태"]],
            use_container_width=True,
            hide_index=True
        )

    if exp_cars_list:
        st.markdown(f"#### 📅 정기검사 만료 임박 차량 ({len(exp_cars_list)}대)")
        for exp_txt in exp_cars_list:
            st.warning(exp_txt)

    if st.button("확인완료", type="primary", use_container_width=True):
        st.rerun()


# ---------------------------------------------------------
# 5. 데이터 불러오기 및 사이드바 구성 (비밀번호 인증 포함)
# ---------------------------------------------------------
(df_cars, df_rules, df_cons, df_reps, df_fbs, df_man), db_status_msg, gs_err_detail = load_all_data()

if "shown_car_popups" not in st.session_state:
    st.session_state.shown_car_popups = set()
if "shown_admin_popup" not in st.session_state:
    st.session_state.shown_admin_popup = False
if "admin_authenticated" not in st.session_state:
    st.session_state.admin_authenticated = False

st.sidebar.title("🚘 차량 QR 관리 시스템")
st.sidebar.caption(db_status_msg)
if gs_err_detail:
    st.sidebar.warning(f"⚠️ 연동 점검: {gs_err_detail}")
if st.sidebar.button("🔄 데이터 새로고침", use_container_width=True):
    st.rerun()

st.sidebar.divider()

query_params = st.query_params
qr_car_param = query_params.get("car", None)

nav_options = ["탑승차량 정보조회", "관리자 대시보드"]
selected_nav = st.sidebar.radio(
    "🖥️ 화면 모드 선택",
    nav_options,
    index=0
)

if selected_nav == "관리자 대시보드":
    if not st.session_state.admin_authenticated:
        st.sidebar.subheader("🔒 관리자 인증 필요")
        pw_input = st.sidebar.text_input("관리자 비밀번호", type="password")
        if st.sidebar.button("로그인", use_container_width=True):
            if pw_input == "kpetro!23":
                st.session_state.admin_authenticated = True
                st.rerun()
            else:
                st.sidebar.error("비밀번호가 올바르지 않습니다.")
        view_mode = "탑승차량 정보조회"
    else:
        view_mode = "관리자 대시보드"
        if st.sidebar.button("관리자 로그아웃", use_container_width=True):
            st.session_state.admin_authenticated = False
            st.rerun()
else:
    view_mode = "탑승차량 정보조회"

st.sidebar.divider()

# ---------------------------------------------------------
# 6. 화면 모드 1: 탑승차량 정보조회 (현장)
# ---------------------------------------------------------
if view_mode == "탑승차량 정보조회":
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
        insp_badge, insp_diff = calc_dday_badge(car_row.get("정기검사 만료일", ""))

        car_cons = df_cons[df_cons["차량명"] == sel_car]
        cons_alerts = []
        for _, c_row in car_cons.iterrows():
            badge, driven_km, m_elapsed, _ = evaluate_consumable_status(
                current_km,
                c_row["마지막 교체 주행거리"],
                c_row["교체주기(Km)"],
                c_row["마지막 교체일"],
                c_row["교체주기(개월)"]
            )
            if "🔴" in badge or "🟡" in badge:
                cons_alerts.append({
                    "항목": c_row["항목"],
                    "badge": badge,
                    "driven_km": driven_km,
                    "m_elapsed": m_elapsed,
                    "cycle_km": safe_int(c_row["교체주기(Km)"]),
                    "cycle_m": safe_int(c_row["교체주기(개월)"])
                })

        car_fbs = df_fbs[df_fbs["차량명"] == sel_car].sort_values(by="날짜", ascending=False)
        fb_alerts = []
        for _, f_row in car_fbs.iterrows():
            if "조치필요" in str(f_row["처리상태"]) or "조치중" in str(f_row["처리상태"]):
                fb_alerts.append(f_row.to_dict())

        insp_alert_msg = None
        if insp_diff <= 30:
            insp_alert_msg = f"정기검사 만료일({car_row.get('정기검사 만료일', '-')})이 {insp_badge} 상태입니다."

        has_any_alert = bool(cons_alerts or fb_alerts or insp_alert_msg)
        if has_any_alert and sel_car not in st.session_state.shown_car_popups:
            st.session_state.shown_car_popups.add(sel_car)
            show_car_alert_dialog(sel_car, cons_alerts, fb_alerts, insp_alert_msg)

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

        c1, c2, c3 = st.columns([1.3, 1.3, 1.4])
        c1.metric("현재 누적 주행거리", f"{current_km:,} km")
        c2.metric("정기검사 만료일", str(car_row.get("정기검사 만료일", "-")), insp_badge)

        with c3:
            with st.popover("📍 주행거리 즉시 업데이트", use_container_width=True):
                new_km_val = st.number_input("계기판 현재 주행거리 (km)", min_value=0, value=current_km, step=100)
                if st.button("주행거리 저장", type="primary", use_container_width=True):
                    df_cars.loc[df_cars["차량명"] == sel_car, "최근 주행거리"] = int(new_km_val)
                    save_all_data(df_cars, df_rules, df_cons, df_reps, df_fbs, df_man)
                    st.toast("주행거리가 업데이트되었습니다!")
                    st.rerun()

            if has_any_alert:
                if st.button("🚨 차량 주의사항 팝업 다시보기", use_container_width=True):
                    show_car_alert_dialog(sel_car, cons_alerts, fb_alerts, insp_alert_msg)

        tab_maint, tab_fb, tab_manual = st.tabs(["🛠️ 정비·소모품 이력", "💬 특이사항", "🚨 사고절차"])

        # [탭 1] 정비·소모품 이력
        with tab_maint:
            st.subheader("🔧 소모품 교체 주기 및 상태 점검")
            if car_cons.empty:
                st.info("등록된 소모품 관리 항목이 없습니다.")
            else:
                cols_cons = st.columns(2)
                for idx, (_, c_row) in enumerate(car_cons.iterrows()):
                    last_km_val = safe_int(c_row["마지막 교체 주행거리"])
                    cycle_km_val = safe_int(c_row["교체주기(Km)"])
                    cycle_m_val = safe_int(c_row["교체주기(개월)"])
                    rec_target_km = last_km_val + cycle_km_val

                    badge, driven_km, m_elapsed, ratio = evaluate_consumable_status(
                        current_km,
                        last_km_val,
                        cycle_km_val,
                        c_row["마지막 교체일"],
                        cycle_m_val
                    )
                    with cols_cons[idx % 2]:
                        st.markdown(f"""
                        <div class="status-card">
                            <b>{c_row['항목']}</b> &nbsp; <span style="float:right; font-weight:bold;">{badge}</span><br>
                            <small style="color:#475569; line-height:1.6;">
                                마지막 교체: {c_row['마지막 교체일']} ({last_km_val:,} km)<br>
                                권장 교체시기: <b>{rec_target_km:,} km</b> ({cycle_m_val}개월 주기)<br>
                                경과: <b>+{driven_km:,} km ({m_elapsed}개월 경과)</b> / 기준 {cycle_km_val:,} km ({cycle_m_val}개월)
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
                    
                    all_rule_items = df_rules["항목"].tolist() if not df_rules.empty else ["엔진오일"]
                    r_item = f_col2.selectbox("연동할 소모품 항목 (소모품 교체 시)", ["선택 안함"] + all_rule_items)
                    
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

                            if r_km > current_km:
                                df_cars.loc[df_cars["차량명"] == sel_car, "최근 주행거리"] = int(r_km)

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
            if car_fbs.empty:
                st.info("접수된 차량 특이사항이 없습니다.")
            else:
                st.dataframe(car_fbs, use_container_width=True, hide_index=True)

            st.markdown("#### ✍️ 운행 중 특이사항 신규 등록")
            with st.form("add_feedback_form", clear_on_submit=True):
                fb_c1, fb_c2 = st.columns(2)
                fb_writer = fb_c1.text_input("작성자 (성명/부서명)")
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
                            "처리상태": "조치필요"
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
# 7. 화면 모드 2: 관리자 대시보드
# ---------------------------------------------------------
else:
    branch_list = ["전체 보기"] + sorted(df_cars["소속본부"].dropna().unique().tolist())
    selected_branch = st.sidebar.selectbox("조회 대상 본부", branch_list)

    filtered_cars = df_cars if selected_branch == "전체 보기" else df_cars[df_cars["소속본부"] == selected_branch]
    target_car_names = filtered_cars["차량명"].tolist()

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

    filtered_fbs = df_fbs[df_fbs["차량명"].isin(target_car_names)]
    pending_fbs_df = filtered_fbs[filtered_fbs["처리상태"].astype(str).str.contains("조치필요|조치중")]

    exp_cars_list = []
    for _, car_r in filtered_cars.iterrows():
        b_txt, d1 = calc_dday_badge(car_r.get("정기검사 만료일", ""))
        if d1 <= 30:
            exp_cars_list.append(f"{car_r['차량명']} ({car_r['차종']}) - 정기검사 만료일: {car_r.get('정기검사 만료일', '-')} [{b_txt}]")

    has_admin_alerts = bool(alert_cons_list or not pending_fbs_df.empty or exp_cars_list)
    if has_admin_alerts and not st.session_state.shown_admin_popup:
        st.session_state.shown_admin_popup = True
        show_admin_alert_dialog(selected_branch, alert_cons_list, pending_fbs_df, exp_cars_list)

    title_col1, title_col2 = st.columns([3.5, 1])
    with title_col1:
        st.title("📊 전사 차량 QR 이력·상태 통합 관제 대시보드")
    with title_col2:
        st.write("")
        if has_admin_alerts:
            if st.button("전체 차량 상태정보 알림 확인", use_container_width=True):
                show_admin_alert_dialog(selected_branch, alert_cons_list, pending_fbs_df, exp_cars_list)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("🚘 관리 대상 차량", f"{len(filtered_cars)} 대")
    k2.metric("🔧 소모품 점검·교체 도래", f"{len(alert_cons_list)} 건")
    k3.metric("💬 미조치 현장 특이사항", f"{len(pending_fbs_df)} 건")
    k4.metric("📅 정기검사 만료 임박(30일내)", f"{len(exp_cars_list)} 대")

    st.divider()

    adm_tab1, adm_tab2, adm_tab3, adm_tab4 = st.tabs([
        "전체 차량 현황 및 정비이력 모아보기",
        "현장 접수 특이사항 조치현황",
        "전체 차량 데이터 수정",
        "차량 부착용 QR 코드 생성기"
    ])

    # [관리자 탭 1] 전체 차량 현황 및 정비이력 모아보기
    with adm_tab1:
        st.subheader("전체 업무용 차량 현황")
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

    # [관리자 탭 2] 현장 접수 특이사항 조치현황
    with adm_tab2:
        st.subheader("현장 접수 특이사항 조치현황")
        st.caption("표에서 '처리상태'와 '확인일자'를 직접 수정한 뒤 아래 저장 버튼을 누르면 즉시 반영됩니다.")
        
        edited_fbs = st.data_editor(
            df_fbs,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "처리상태": st.column_config.SelectboxColumn(
                    "처리상태",
                    options=["조치필요", "조치중", "조치완료"],
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

    # [관리자 탭 3] 전체 차량 데이터 수정
    with adm_tab3:
        st.subheader("전체 차량 데이터 수정")
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

    # [관리자 탭 4] 차량 부착용 QR 코드 생성기
    with adm_tab4:
        st.subheader("차량 부착용 QR 코드 생성기")
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