import streamlit as st
import datetime
import calendar
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 1. 페이지 기본 설정 및 반응형 그리드 CSS (높이 동일 맞춤)
st.set_page_config(page_title="9BLOCK 통합 오픈 스케줄러", page_icon="🗓️", layout="wide")

st.markdown("""
    <style>
    /* PC 달력: 같은 주의 모든 셀 높이가 동적으로 같이 늘어나는 Flex 그리드 */
    .pc-week-row {
        display: flex;
        width: 100%;
        gap: 4px;
        margin-bottom: 4px;
    }
    .pc-day-cell {
        flex: 1;
        min-width: 0;
        min-height: 120px;
        border: 1px solid #e0e0e0;
        padding: 6px;
        border-radius: 6px;
        background-color: #ffffff;
        box-sizing: border-box;
    }
    .pc-day-cell.other-month {
        background-color: #f9f9f9;
        color: #bbb;
    }
    
    /* 헤더 그리드 */
    .pc-header-row {
        display: flex;
        width: 100%;
        gap: 4px;
        margin-bottom: 6px;
        text-align: center;
        font-weight: bold;
        background-color: #1F4E78;
        color: white;
        padding: 8px 0;
        border-radius: 6px;
    }
    .pc-header-cell {
        flex: 1;
    }

    /* 지점 구분 (직영/가맹/위탁) 전용 뱃지 */
    .badge-type-가맹 { background-color: #e3f2fd; color: #0d47a1; padding: 1px 4px; border-radius: 3px; font-size: 10px; font-weight: bold; }
    .badge-type-직영 { background-color: #e8f5e9; color: #1b5e20; padding: 1px 4px; border-radius: 3px; font-size: 10px; font-weight: bold; }
    .badge-type-위탁 { background-color: #fff3e0; color: #e65100; padding: 1px 4px; border-radius: 3px; font-size: 10px; font-weight: bold; }

    /* 모바일 달력 그리드 */
    .mobile-cal-header {
        display: grid;
        grid-template-columns: repeat(7, 1fr);
        text-align: center;
        font-weight: bold;
        background-color: #1F4E78;
        color: white;
        padding: 6px 0;
        border-radius: 6px 6px 0 0;
        font-size: 12px;
    }
    .mobile-cal-grid {
        display: grid;
        grid-template-columns: repeat(7, 1fr);
        gap: 2px;
        background-color: #e0e0e0;
        border: 1px solid #e0e0e0;
        border-radius: 0 0 6px 6px;
    }
    .mobile-day-cell {
        background-color: #ffffff;
        min-height: 100px;
        padding: 3px;
        font-size: 10px;
    }
    .mobile-day-cell.other-month {
        background-color: #f7f7f7;
        color: #bbb;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🗓️ 9BLOCK 통합 오픈 스케줄러")

# 기본 템플릿 정의
DEFAULT_STORES = {
    "충청점": {"date": datetime.date(2026, 9, 18), "type": "가맹"}
}

DEFAULT_TASKS = [
    {"부서": "개발", "주요업무": "백화점/몰 입점 물건조사 및 상권분석", "offset": -45, "담당": "개발팀"},
    {"부서": "구매/운영", "주요업무": "[★선행] 입지 맞춤 메뉴 라인업 확정", "offset": -40, "담당": "구매/운영팀"},
    {"부서": "개발", "주요업무": "가맹점주 미팅 및 전대차 계약 체결", "offset": -35, "담당": "개발팀"},
    {"부서": "인테리어", "주요업무": "확정 장비 스펙 기반 예가 산출", "offset": -33, "담당": "인테리어팀"},
    {"부서": "인테리어", "주요업무": "백화점 도면/VMD 제출 및 도면 승인", "offset": -30, "담당": "인테리어팀"},
    {"부서": "마케팅", "주요업무": "특수상권 프로모션 기획 및 백화점 앱 DM 협의", "offset": -20, "담당": "마케팅팀"},
    {"부서": "인테리어", "주요업무": "인테리어 현장 착공 (야간공사 및 주방/전기)", "offset": -16, "담당": "인테리어팀"},
    {"부서": "구매/운영", "주요업무": "점주/매니저 본사 레시피 교육 및 현장 실습", "offset": -10, "담당": "운영팀"},
    {"부서": "마케팅", "주요업무": "POP/POS 홍보물 제작 및 매장 입고/설치", "offset": -7, "담당": "마케팅팀"},
    {"부서": "전부서", "주요업무": "🚨 [오픈일사수] 준공검사 및 최종 시운전", "offset": -2, "담당": "전부서"},
    {"부서": "전부서", "주요업무": "🎉 GRAND OPEN & 현장 지원", "offset": 0, "담당": "전부서"}
]

DEFAULT_CONTACTS = {
    "개발팀": {"name": "김개발 팀장", "email": "dev@9block.co.kr"},
    "구매/운영팀": {"name": "이구매 팀장", "email": "buy@9block.co.kr"},
    "인테리어팀": {"name": "박설계 팀장", "email": "interior@9block.co.kr"},
    "마케팅팀": {"name": "최홍보 팀장", "email": "mkt@9block.co.kr"},
    "운영팀": {"name": "정운영 팀장", "email": "ops@9block.co.kr"},
    "전부서": {"name": "오픈지원TF", "email": "tf@9block.co.kr"}
}

# --- 구글 시트 연동 ---
def get_gsheet_doc():
    if "gcp_service_account" not in st.secrets or "SPREADSHEET_ID" not in st.secrets:
        return None
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials_info = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_info, scope)
        client = gspread.authorize(creds)
        sheet_id = st.secrets["SPREADSHEET_ID"]
        return client.open_by_key(sheet_id)
    except Exception as e:
        st.error(f"구글 시트 연동 실패: {e}")
        return None

def load_data_from_gsheets():
    try:
        doc = get_gsheet_doc()
        if not doc: return dict(DEFAULT_STORES), dict(DEFAULT_CONTACTS), {}

        # 1. 지점
        stores = {}
        try:
            ws_stores = doc.worksheet("stores")
            records = ws_stores.get_all_records()
            for r in records:
                s_name = str(r.get("매장명", "")).strip()
                s_date = str(r.get("오픈일", "")).strip()
                s_type = str(r.get("구분", "가맹")).strip()
                if s_name and s_date:
                    stores[s_name] = {
                        "date": datetime.datetime.strptime(s_date, "%Y-%m-%d").date(),
                        "type": s_type if s_type in ["직영", "위탁", "가맹"] else "가맹"
                    }
        except Exception: pass
        if not stores: stores = dict(DEFAULT_STORES)

        # 2. 주소록
        contacts = {}
        try:
            ws_contacts = doc.worksheet("contacts")
            c_records = ws_contacts.get_all_records()
            for r in c_records:
                team = str(r.get("팀명", "")).strip()
                name = str(r.get("담당자", "")).strip()
                email = str(r.get("이메일", "")).strip()
                if team: contacts[team] = {"name": name, "email": email}
        except Exception: pass
        if not contacts: contacts = dict(DEFAULT_CONTACTS)

        # 3. 완료 상태 및 커스텀 설정
        completed_tasks = {}
        try:
            ws_status = doc.worksheet("task_status")
            s_records = ws_status.get_all_records()
            for r in s_records:
                t_id = str(r.get("task_id", "")).strip()
                if t_id:
                    completed_tasks[t_id] = {
                        "completed": bool(r.get("completed", False)),
                        "custom_task": str(r.get("custom_task", "")),
                        "custom_offset": int(r.get("custom_offset", 0)) if str(r.get("custom_offset", "")).strip() != "" else None,
                        "custom_assignee": str(r.get("custom_assignee", ""))
                    }
        except Exception: pass

        return stores, contacts, completed_tasks
    except Exception:
        return dict(DEFAULT_STORES), dict(DEFAULT_CONTACTS), {}

def save_stores_to_gsheets():
    doc = get_gsheet_doc()
    if not doc: return
    try:
        try: ws = doc.worksheet("stores")
        except: ws = doc.add_worksheet(title="stores", rows="100", cols="5")
        ws.clear()
        rows = [["매장명", "오픈일", "구분"]]
        for name, info in st.session_state.stores.items():
            rows.append([name, info["date"].strftime("%Y-%m-%d"), info.get("type", "가맹")])
        ws.update("A1", rows)
        st.toast("☁️ 지점 정보가 저장되었습니다!", icon="✅")
    except Exception as e: st.error(f"저장 실패: {e}")

def save_contacts_to_gsheets():
    doc = get_gsheet_doc()
    if not doc: return
    try:
        try: ws = doc.worksheet("contacts")
        except: ws = doc.add_worksheet(title="contacts", rows="100", cols="5")
        ws.clear()
        rows = [["팀명", "담당자", "이메일"]]
        for team, info in st.session_state.contacts.items():
            rows.append([team, info["name"], info["email"]])
        ws.update("A1", rows)
        st.toast("☁️ 주소록이 저장되었습니다!", icon="✅")
    except Exception as e: st.error(f"저장 실패: {e}")

def save_task_status_to_gsheets():
    doc = get_gsheet_doc()
    if not doc: return
    try:
        try: ws = doc.worksheet("task_status")
        except: ws = doc.add_worksheet(title="task_status", rows="500", cols="6")
        ws.clear()
        rows = [["task_id", "completed", "custom_task", "custom_offset", "custom_assignee"]]
        for t_id, info in st.session_state.task_status.items():
            c_val = "TRUE" if info.get("completed") else "FALSE"
            rows.append([
                t_id, c_val, 
                info.get("custom_task", ""), 
                info.get("custom_offset", ""), 
                info.get("custom_assignee", "")
            ])
        ws.update("A1", rows)
        st.toast("☁️ 공정 상태가 구글 시트에 즉시 반영되었습니다!", icon="✅")
    except Exception as e: st.error(f"상태 저장 실패: {e}")

# 세션 초기화
if "stores" not in st.session_state or "contacts" not in st.session_state:
    loaded_stores, loaded_contacts, loaded_status = load_data_from_gsheets()
    st.session_state.stores = loaded_stores
    st.session_state.contacts = loaded_contacts
    st.session_state.task_status = loaded_status

if "master_tasks" not in st.session_state:
    st.session_state.master_tasks = list(DEFAULT_TASKS)

if "current_view_date" not in st.session_state:
    st.session_state.current_view_date = datetime.date(2026, 8, 1)

if "view_mode_choice" not in st.session_state:
    st.session_state.view_mode_choice = "🖥️ PC 넓은 달력 보기"

# --- 메일 발송 함수 (전체 발송 보완) ---
def send_email_auto(receiver_emails, subject, body):
    try:
        if "SENDER_EMAIL" not in st.secrets or "SENDER_PASSWORD" not in st.secrets:
            return False, "Streamlit Secrets에 SENDER_EMAIL/SENDER_PASSWORD가 설정되지 않았습니다."
        sender_email = st.secrets["SENDER_EMAIL"]
        sender_password = st.secrets["SENDER_PASSWORD"]

        if isinstance(receiver_emails, str):
            receiver_emails = [receiver_emails]
        
        valid_receivers = [e.strip() for e in receiver_emails if e and "@" in str(e)]
        if not valid_receivers:
            return False, "수신 가능한 이메일 주소가 없습니다."

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = ", ".join(valid_receivers)
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, valid_receivers, msg.as_string())
        server.quit()
        return True, f"총 {len(valid_receivers)}명에게 이메일 알림 발송 성공"
    except Exception as e:
        return False, f"발송 실패: {e}"

def notify_all_contacts(subject, body):
    all_emails = [info["email"] for info in st.session_state.contacts.values() if info.get("email")]
    return send_email_auto(all_emails, subject, body)

def get_dept_color(dept_name):
    color_map = {
        "개발": {"bg": "#e3f2fd", "text": "#1565c0", "border": "#90caf9"},
        "구매/운영": {"bg": "#e8f5e9", "text": "#2e7d32", "border": "#a5d6a7"},
        "인테리어": {"bg": "#fff3e0", "text": "#e65100", "border": "#ffcc80"},
        "마케팅": {"bg": "#f3e5f5", "text": "#7b1fa2", "border": "#ce93d8"},
        "전부서": {"bg": "#ffebee", "text": "#c62828", "border": "#ef9a9a"}
    }
    return color_map.get(dept_name, {"bg": "#f5f5f5", "text": "#424242", "border": "#e0e0e0"})

# --- 사이드바 메뉴 ---
st.sidebar.header("➕ 신규 지점 등록")
with st.sidebar.form("add_store_form", clear_on_submit=True):
    new_store_name = st.text_input("매장명 (예: 강남점)")
    new_store_type = st.selectbox("지점 구분", ["가맹", "직영", "위탁"])
    new_open_date = st.date_input("GRAND OPEN 예정일", value=datetime.date(2026, 10, 15))
    if st.form_submit_button("지점 추가하기"):
        if new_store_name.strip():
            st.session_state.stores[new_store_name.strip()] = {
                "date": new_open_date,
                "type": new_store_type
            }
            save_stores_to_gsheets()
            
            sub = f"📢 [신규 지점 추가] '{new_store_name.strip()}'({new_store_type}) 오픈 스케줄 등록"
            msg = f"안녕하세요,\n\n신규 지점 [{new_store_name.strip()}] ({new_store_type})의 오픈 일정이 등록되었습니다.\n오픈 예정일: {new_open_date}\n\n스케줄러 앱에서 상세 공정을 확인해 주세요."
            notify_all_contacts(sub, msg)

            st.sidebar.success(f"'{new_store_name}' 추가 및 전체 담당자 메일 발송 완료!")
            st.rerun()

st.sidebar.markdown("---")
st.sidebar.header("✏️ 지점 정보 수정 / 삭제")
if st.session_state.stores:
    store_list = list(st.session_state.stores.keys())
    selected_edit_store = st.sidebar.selectbox("지점 선택", store_list)
    
    curr_store_info = st.session_state.stores[selected_edit_store]
    current_date = curr_store_info["date"]
    current_type = curr_store_info.get("type", "가맹")
    
    updated_type = st.sidebar.selectbox("지점 구분 변경", ["가맹", "직영", "위탁"], index=["가맹", "직영", "위탁"].index(current_type), key=f"type_select_{selected_edit_store}")
    updated_date = st.sidebar.date_input("오픈 예정일 변경", value=current_date, key=f"date_input_{selected_edit_store}")
    
    col_btn1, col_btn2 = st.sidebar.columns(2)
    if col_btn1.button("정보 수정"):
        old_date = current_date
        st.session_state.stores[selected_edit_store] = {
            "date": updated_date,
            "type": updated_type
        }
        save_stores_to_gsheets()
        
        sub = f"🚨 [오픈일 변경 알림] [{selected_edit_store}] 지점 오픈 정보가 수정되었습니다."
        msg = f"안녕하세요,\n\n[{selected_edit_store}] 지점의 기본 정보가 변경되었습니다.\n\n"
        msg += f"• 구분: {current_type} ➡️ {updated_type}\n"
        msg += f"• 오픈 예정일: {old_date} ➡️ {updated_date}\n\n"
        msg += f"모든 공정 D-Day 스케줄이 자동으로 재조정되었습니다."
        
        notify_all_contacts(sub, msg)
        st.sidebar.success("수정 완료 및 전체 알림 발송!")
        st.rerun()

    if col_btn2.button("지점 삭제"):
        del st.session_state.stores[selected_edit_store]
        save_stores_to_gsheets()
        st.sidebar.warning("삭제 완료!")
        st.rerun()

# --- 데이터 연산 ---
all_schedule_data = []

for s_name, s_info in st.session_state.stores.items():
    s_open_date = s_info["date"]
    s_type = s_info.get("type", "가맹")
    
    for task in st.session_state.master_tasks:
        task_id = f"{s_name}_{task['부서']}_{task['주요업무']}"
        status_info = st.session_state.task_status.get(task_id, {})
        
        task_title = status_info.get("custom_task") if status_info.get("custom_task") else task["주요업무"]
        offset_val = status_info.get("custom_offset") if status_info.get("custom_offset") is not None else task["offset"]
        assignee_val = status_info.get("custom_assignee") if status_info.get("custom_assignee") else task["담당"]
        is_completed = status_info.get("completed", False)

        task_date = s_open_date + datetime.timedelta(days=offset_val)
        date_str = task_date.strftime("%Y-%m-%d")
        d_day_str = f"D{offset_val}" if offset_val < 0 else ("D-Day" if offset_val == 0 else f"D+{offset_val}")

        all_schedule_data.append({
            "task_id": task_id,
            "일자": date_str, "D-Day": d_day_str, "매장명": s_name, "지점타입": s_type,
            "부서": task["부서"], "주요업무": task_title, "담당자": assignee_val,
            "raw_date": task_date, "year": task_date.year, "month": task_date.month,
            "completed": is_completed, "offset": offset_val
        })

# 메인 탭
tab1, tab2, tab3, tab4 = st.tabs(["🗓️ 월별 달력", "📋 지점별 공정표 수정", "📮 자동 발송 현황", "👤 이메일 주소록"])

# --- TAB 1: 월별 달력 ---
with tab1:
    c_filter1, c_filter2, c_filter3 = st.columns(3)
    selected_store_filter = c_filter1.selectbox("🏬 지점 선택 필터", ["전체 지점"] + list(st.session_state.stores.keys()))
    type_filter = c_filter2.selectbox("🏷️ 지점 구분 필터", ["전체 구분", "직영", "위탁", "가맹"])
    status_filter = c_filter3.radio("📌 진행 상태 필터", ["전체 보기", "진행 예정만", "완료된 항목만"], horizontal=True)

    st.markdown("---")
    c_prev, c_title, c_next = st.columns([1, 4, 1])
    if c_prev.button("◀ 이전달", use_container_width=True):
        y, m = st.session_state.current_view_date.year, st.session_state.current_view_date.month - 1
        if m == 0: y, m = y - 1, 12
        st.session_state.current_view_date = datetime.date(y, m, 1)
        st.rerun()
    if c_next.button("다음달 ▶", use_container_width=True):
        y, m = st.session_state.current_view_date.year, st.session_state.current_view_date.month + 1
        if m == 13: y, m = y + 1, 1
        st.session_state.current_view_date = datetime.date(y, m, 1)
        st.rerun()

    v_year, v_month = st.session_state.current_view_date.year, st.session_state.current_view_date.month
    c_title.markdown(f"<h3 style='text-align: center; color: #1F4E78;'>🗓️ {v_year}년 {v_month:02d}월 가맹점 스케줄</h3>", unsafe_allow_html=True)

    # 필터링
    filtered_schedules = []
    for item in all_schedule_data:
        if selected_store_filter != "전체 지점" and item["매장명"] != selected_store_filter:
            continue
        if type_filter != "전체 구분" and item["지점타입"] != type_filter:
            continue
        if status_filter == "진행 예정만" and item["completed"]:
            continue
        if status_filter == "완료된 항목만" and not item["completed"]:
            continue
        filtered_schedules.append(item)

    schedule_map = {}
    for item in filtered_schedules:
        d_str = item["일자"]
        if d_str not in schedule_map: schedule_map[d_str] = []
        schedule_map[d_str].append(item)

    options = ["🖥️ PC 넓은 달력 보기", "📱 모바일 달력 보기"]
    current_index = options.index(st.session_state.view_mode_choice) if st.session_state.view_mode_choice in options else 0
    selected_mode = st.radio("🖥️ 화면 보기 방식 선택", options, index=current_index, horizontal=True)
    st.session_state.view_mode_choice = selected_mode

    # 1. PC 넓은 달력 보기 (Flex 동적 높이 + 즉시 체크 기능 반영)
    if st.session_state.view_mode_choice == "🖥️ PC 넓은 달력 보기":
        cal = calendar.Calendar(firstweekday=0)
        month_days = cal.monthdatescalendar(v_year, v_month)
        
        # 헤더
        header_html = "<div class='pc-header-row'>"
        for d_name in ["월", "화", "수", "목", "금", "토", "일"]:
            header_html += f"<div class='pc-header-cell'>{d_name}</div>"
        header_html += "</div>"
        st.markdown(header_html, unsafe_allow_html=True)
        
        # 주 단위 루프
        for week in month_days:
            cols = st.columns(7)
            for i, d in enumerate(week):
                d_str = d.strftime("%Y-%m-%d")
                is_cur_month = (d.month == v_month)
                
                with cols[i]:
                    # 날짜 헤더 표기
                    bg_col = "#ffffff" if is_cur_month else "#f9f9f9"
                    text_col = "#000000" if is_cur_month else "#bbbbbb"
                    st.markdown(f"<div style='background-color:{bg_col}; color:{text_col}; font-weight:bold; padding:2px 4px; border-radius:4px;'>{d.day}일</div>", unsafe_allow_html=True)
                    
                    if d_str in schedule_map:
                        for item in schedule_map[d_str]:
                            t_id = item["task_id"]
                            c_info = get_dept_color(item["부서"])
                            
                            # 달력 내부 인터렉티브 완료 체크 박스
                            checked = st.checkbox(
                                f"[{item['지점타입']}|{item['매장명']}] {item['부서']}",
                                value=item["completed"],
                                key=f"cal_check_{t_id}"
                            )
                            
                            # 체크 상태 변경 감지 및 즉시 반영 & 전체 메일 발송
                            if checked != item["completed"]:
                                if t_id not in st.session_state.task_status:
                                    st.session_state.task_status[t_id] = {}
                                st.session_state.task_status[t_id]["completed"] = checked
                                save_task_status_to_gsheets()
                                
                                # 전체 알림 발송
                                state_str = "완료 (✅)" if checked else "진행 예정 (⏳)"
                                sub = f"📢 [공정 상태 변경] [{item['매장명']}] {item['부서']} - {item['주요업무']}"
                                body = f"안녕하세요,\n\n[{item['매장명']}] 지점의 공정 상태가 변경되었습니다.\n\n"
                                body += f"📌 공정명: {item['주요업무']}\n📌 담당부서: {item['부서']}\n📌 상태: {state_str}\n\n스케줄러 앱에서 확인 가능합니다."
                                
                                notify_all_contacts(sub, body)
                                st.rerun()

                            # 공정 내용 표시 카드
                            text_style = "text-decoration: line-through; color: #888;" if item["completed"] else "color: #222;"
                            st.markdown(f"""
                            <div style='margin-bottom:6px; padding:4px; border-radius:4px; background-color:{c_info["bg"]}; border:1px solid {c_info["border"]};'>
                                <div style='font-size:11px; {text_style}; line-height:1.2;'>{item["주요업무"]}</div>
                            </div>
                            """, unsafe_allow_html=True)

    # 2. 모바일 달력 보기
    else:
        cal = calendar.Calendar(firstweekday=0)
        month_days = cal.monthdatescalendar(v_year, v_month)
        
        cal_html = "<div class='mobile-cal-header'>"
        for d_name in ["월", "화", "수", "목", "금", "토", "일"]:
            cal_html += f"<div>{d_name}</div>"
        cal_html += "</div>"
        
        cal_html += "<div class='mobile-cal-grid'>"
        for week in month_days:
            for d in week:
                d_str = d.strftime("%Y-%m-%d")
                is_cur_month = (d.month == v_month)
                cell_class = "mobile-day-cell" if is_cur_month else "mobile-day-cell other-month"
                
                tasks_html = ""
                if d_str in schedule_map:
                    for item in schedule_map[d_str]:
                        c_info = get_dept_color(item["부서"])
                        status_mark = "✅" if item["completed"] else ""
                        bg_color = "#e0e0e0" if item["completed"] else c_info["bg"]
                        
                        tasks_html += f"""
                        <div style='margin-top:2px; padding:2px; border-radius:3px; background-color:{bg_color}; border:1px solid {c_info["border"]}; font-size:9px;'>
                            {status_mark}<b>[{item["지점타입"][:1]}|{item["매장명"][:2]}]</b> [{item["부서"][:2]}] {item["주요업무"][:5]}
                        </div>"""
                
                cal_html += f"<div class='{cell_class}'><b>{d.day}</b>{tasks_html}</div>"
        cal_html += "</div>"
        
        st.markdown(cal_html, unsafe_allow_html=True)

# --- TAB 2: 지점별 공정표 수정 ---
with tab2:
    st.subheader("📋 지점별 공정표 확인 / 수정 / 담당자 지정")
    
    if st.session_state.stores:
        store_tabs = st.tabs([f"[{info.get('type','가맹')}] {sname}" for sname, info in st.session_state.stores.items()])
        
        for index, (s_name, s_info) in enumerate(st.session_state.stores.items()):
            with store_tabs[index]:
                st.markdown(f"#### 🏬 {s_name} [{s_info.get('type','가맹')}] 세부 관리 (오픈 예정일: {s_info['date']})")
                
                store_tasks = [t for t in all_schedule_data if t["매장명"] == s_name]
                
                for task in store_tasks:
                    t_id = task["task_id"]
                    
                    with st.expander(f"[{task['D-Day']}] {task['일자']} | [{task['부서']}] {task['주요업무']} (담당: {task['담당자']}) {'(✅ 완료됨)' if task['completed'] else ''}"):
                        col1, col2 = st.columns([2, 1])
                        
                        is_done = col1.checkbox("✅ 공정 완료 처리", value=task["completed"], key=f"tab2_check_{t_id}")
                        new_title = col1.text_input("공정명 수정", value=task["주요업무"], key=f"title_{t_id}")
                        new_offset = col1.number_input("D-Day 설정 (오픈일 기준 일수, 예: -45)", value=task["offset"], key=f"offset_{t_id}")
                        
                        existing_teams = list(st.session_state.contacts.keys())
                        curr_assignee_idx = existing_teams.index(task["담당자"]) if task["담당자"] in existing_teams else 0
                        selected_assignee = col1.selectbox("담당팀/담당자 지정", existing_teams, index=curr_assignee_idx, key=f"assignee_select_{t_id}")
                        
                        col2.markdown("**➕ 신규 담당자 즉시 추가**")
                        new_team_input = col2.text_input("신규 팀명/구분", key=f"new_team_{t_id}")
                        new_name_input = col2.text_input("담당자 성함", key=f"new_name_{t_id}")
                        new_email_input = col2.text_input("이메일 주소", key=f"new_email_{t_id}")
                        
                        if col2.button("담당자 주소록에 추가", key=f"add_contact_btn_{t_id}"):
                            if new_team_input.strip() and new_email_input.strip():
                                st.session_state.contacts[new_team_input.strip()] = {
                                    "name": new_name_input.strip(),
                                    "email": new_email_input.strip()
                                }
                                save_contacts_to_gsheets()
                                st.success("신규 담당자 등록 성공!")
                                st.rerun()
                            else:
                                col2.error("팀명과 이메일은 필수 입력 항목입니다.")

                        st.markdown("---")
                        if st.button("💾 변경사항 저장 및 전체 담당자 메일 발송", key=f"save_{t_id}"):
                            if t_id not in st.session_state.task_status:
                                st.session_state.task_status[t_id] = {}
                            
                            st.session_state.task_status[t_id]["completed"] = is_done
                            st.session_state.task_status[t_id]["custom_task"] = new_title.strip()
                            st.session_state.task_status[t_id]["custom_offset"] = new_offset
                            st.session_state.task_status[t_id]["custom_assignee"] = selected_assignee
                            
                            save_task_status_to_gsheets()
                            
                            subject = f"📢 [공정 변경 알림] [{s_name}] {task['부서']} 공정 수정"
                            body = f"안녕하세요,\n\n[{s_name}] ({s_info.get('type','가맹')}) 지점의 공정 정보가 업데이트되었습니다.\n\n"
                            body += f"📌 공정명: {new_title}\n"
                            body += f"📌 담당: {selected_assignee}\n"
                            body += f"📌 변경일자: D{new_offset} ({s_info['date'] + datetime.timedelta(days=new_offset)})\n"
                            body += f"📌 완료여부: {'완료 (✅)' if is_done else '진행 예정 (⏳)'}\n\n"
                            body += f"스케줄러 앱에서 상세 내용을 확인해 주세요."
                            
                            success, msg = notify_all_contacts(subject, body)
                            if success:
                                st.success(f"저장 완료! 등록된 모든 담당자에게 메일을 발송했습니다. ({msg})")
                            else:
                                st.warning(f"저장 완료 (메일 발송 오류: {msg})")
                            
                            st.rerun()

# --- TAB 3: 자동 발송 현황 ---
with tab3:
    st.subheader("🤖 자동 알림 이메일 발송 현황")
    st.info("💡 매일 아침 8시(한국 시간)에 지정된 공정 담당자에게 알림 메일이 자동 발송됩니다.")

    if all_schedule_data:
        sorted_schedule = sorted(all_schedule_data, key=lambda x: x["일자"])
        for item in sorted_schedule:
            send_target_date = item["raw_date"] - datetime.timedelta(days=1)
            dept_contact = st.session_state.contacts.get(item["담당자"], {"name": "미정", "email": "미등록"})
            
            with st.expander(f"📌 [{item['지점타입']}|{item['매장명']}] [{item['부서']}] {item['주요업무']} (공정일: {item['일자']})"):
                st.write(f"**지정 담당자/팀**: {item['담당자']} ({dept_contact['name']}) | `{dept_contact['email']}`")
                st.write(f"📅 **자동 발송 예정일**: `{send_target_date.strftime('%Y-%m-%d')}` (공정 D-1일전 아침 8시)")

# --- TAB 4: 이메일 주소록 ---
with tab4:
    st.subheader("👤 그룹사 담당자 이메일 주소록 관리")
    c_add1, c_add2, c_add3 = st.columns(3)
    team_input = c_add1.text_input("팀명/구분 (예: 인테리어팀)", key="tab4_team")
    name_input = c_add2.text_input("담당자 성함/직급", value="홍길동 팀장", key="tab4_name")
    email_input = c_add3.text_input("이메일 주소", value="hong@9block.co.kr", key="tab4_email")

    if st.button("➕ 담당자 등록 / 수정", key="tab4_add_btn"):
        if team_input.strip() and email_input.strip():
            st.session_state.contacts[team_input.strip()] = {
                "name": name_input.strip(),
                "email": email_input.strip()
            }
            save_contacts_to_gsheets()
            st.success("등록 완료 및 구글 시트 저장!")
            st.rerun()

    st.markdown("---")
    if st.session_state.contacts:
        for t_name, info in list(st.session_state.contacts.items()):
            st.markdown(f"**{t_name}** | {info['name']} | `{info['email']}`")
