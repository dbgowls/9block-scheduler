import streamlit as st
import datetime
import calendar
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 1. 페이지 기본 설정
st.set_page_config(page_title="9BLOCK 통합 가맹점 오픈 스케줄러", page_icon="🗓️", layout="wide")

st.markdown("""
    <style>
    .mobile-card {
        background-color: #ffffff;
        border-radius: 8px;
        padding: 12px;
        margin-bottom: 10px;
        border-left: 5px solid #1F4E78;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .badge {
        display: inline-block;
        padding: 3px 8px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🗓️ 9BLOCK 가맹점 오픈 스케줄러")

# 기본 템플릿 정의
DEFAULT_STORES = {"충청점": datetime.date(2026, 9, 18)}

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

# --- 구글 시트 클라이언트 연결 ---
def get_gspread_client():
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials_info = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_info, scope)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"구글 연동 키 설정 오류: {e}")
        return None

# 구글 시트 데이터 불러오기
def load_data_from_gsheets():
    client = get_gspread_client()
    if not client:
        return DEFAULT_STORES, DEFAULT_TASKS, DEFAULT_CONTACTS
    
    try:
        sheet_id = st.secrets["SPREADSHEET_ID"]
        doc = client.open_by_key(sheet_id)
        
        # 1) 지점 데이터 불러오기
        try:
            ws_stores = doc.worksheet("stores")
            records = ws_stores.get_all_records()
            stores = {}
            for r in records:
                if r.get("매장명") and r.get("오픈일"):
                    s_date_str = str(r["오픈일"]).strip()
                    stores[str(r["매장명"]).strip()] = datetime.datetime.strptime(s_date_str, "%Y-%m-%d").date()
            if not stores:
                stores = dict(DEFAULT_STORES)
        except Exception:
            stores = dict(DEFAULT_STORES)

        # 2) 공정 데이터
        tasks = list(DEFAULT_TASKS)

        # 3) 주소록 데이터
        try:
            ws_contacts = doc.worksheet("contacts")
            c_records = ws_contacts.get_all_records()
            contacts = {}
            for r in c_records:
                if r.get("팀명"):
                    contacts[str(r["팀명"]).strip()] = {"name": str(r.get("담당자", "")), "email": str(r.get("이메일", ""))}
            if not contacts:
                contacts = dict(DEFAULT_CONTACTS)
        except Exception:
            contacts = dict(DEFAULT_CONTACTS)

        return stores, tasks, contacts

    except Exception as e:
        st.error(f"구글 시트 데이터를 불러오지 못했습니다: {e}")
        return DEFAULT_STORES, DEFAULT_TASKS, DEFAULT_CONTACTS

# 구글 시트에 지점 데이터 저장
def save_stores_to_gsheets():
    client = get_gspread_client()
    if client:
        try:
            sheet_id = st.secrets["SPREADSHEET_ID"]
            doc = client.open_by_key(sheet_id)
            try:
                ws = doc.worksheet("stores")
            except:
                ws = doc.add_worksheet(title="stores", rows="100", cols="5")
            
            ws.clear()
            rows = [["매장명", "오픈일"]]
            for name, date_obj in st.session_state.stores.items():
                rows.append([name, date_obj.strftime("%Y-%m-%d")])
            ws.update("A1", rows)
        except Exception as e:
            st.error(f"구글 시트 저장 실패: {e}")

# 구글 시트에 주소록 데이터 저장
def save_contacts_to_gsheets():
    client = get_gspread_client()
    if client:
        try:
            sheet_id = st.secrets["SPREADSHEET_ID"]
            doc = client.open_by_key(sheet_id)
            try:
                ws = doc.worksheet("contacts")
            except:
                ws = doc.add_worksheet(title="contacts", rows="100", cols="5")
            
            ws.clear()
            rows = [["팀명", "담당자", "이메일"]]
            for team, info in st.session_state.contacts.items():
                rows.append([team, info["name"], info["email"]])
            ws.update("A1", rows)
        except Exception as e:
            st.error(f"구글 시트 주소록 저장 실패: {e}")

# --- 세션 상태 초기화 (구글 시트 연동 갱신) ---
if "stores" not in st.session_state:
    loaded_stores, loaded_tasks, loaded_contacts = load_data_from_gsheets()
    st.session_state.stores = loaded_stores
    st.session_state.master_tasks = loaded_tasks
    st.session_state.contacts = loaded_contacts

if "mail_schedules" not in st.session_state:
    st.session_state.mail_schedules = {}

if "current_view_date" not in st.session_state:
    st.session_state.current_view_date = datetime.date(2026, 8, 1)

if "view_mode_choice" not in st.session_state:
    st.session_state.view_mode_choice = "🗓️ PC용 넓은 달력 보기"

# 이메일 발송 함수
def send_email_notification(sender_email, sender_password, receiver_email, subject, body):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, msg.as_string())
        server.quit()
        return True, "메일 발송에 성공했습니다!"
    except Exception as e:
        return False, f"발송 실패: {e}"

def get_dept_color(dept_name):
    color_map = {
        "개발": {"bg": "#e3f2fd", "text": "#1e88e5", "border": "#90caf9"},
        "구매/운영": {"bg": "#e8f5e9", "text": "#2e7d32", "border": "#a5d6a7"},
        "인테리어": {"bg": "#fff3e0", "text": "#e65100", "border": "#ffcc80"},
        "마케팅": {"bg": "#f3e5f5", "text": "#8e24aa", "border": "#ce93d8"},
        "전부서": {"bg": "#ffebee", "text": "#c62828", "border": "#ef9a9a"}
    }
    return color_map.get(dept_name, {"bg": "#f5f5f5", "text": "#424242", "border": "#e0e0e0"})

# --- 사이드바 메뉴 ---
st.sidebar.header("➕ 신규 지점 등록")
with st.sidebar.form("add_store_form", clear_on_submit=True):
    new_store_name = st.text_input("매장명 (예: 강남점)")
    new_open_date = st.date_input("GRAND OPEN 예정일", value=datetime.date(2026, 10, 15))
    if st.form_submit_button("지점 추가하기"):
        if new_store_name.strip():
            st.session_state.stores[new_store_name.strip()] = new_open_date
            save_stores_to_gsheets()
            st.sidebar.success(f"'{new_store_name}' 추가 완료!")
            st.rerun()

st.sidebar.markdown("---")
st.sidebar.header("✏️ 지점 정보 수정 / 삭제")
if st.session_state.stores:
    selected_edit_store = st.sidebar.selectbox("지점 선택", list(st.session_state.stores.keys()))
    current_date = st.session_state.stores[selected_edit_store]
    updated_date = st.sidebar.date_input("오픈 예정일 변경", value=current_date, key="edit_date_input")
    col_btn1, col_btn2 = st.sidebar.columns(2)
    if col_btn1.button("날짜 수정"):
        st.session_state.stores[selected_edit_store] = updated_date
        save_stores_to_gsheets()
        st.sidebar.success("수정 완료!")
        st.rerun()
    if col_btn2.button("지점 삭제"):
        del st.session_state.stores[selected_edit_store]
        save_stores_to_gsheets()
        st.sidebar.warning("삭제 완료!")
        st.rerun()

st.sidebar.markdown("---")
if st.sidebar.button("🔄 구글 시트 강제 동기화 (새로고침)"):
    del st.session_state.stores
    st.rerun()

# 3. 데이터 연산
all_schedule_data = []
schedule_map = {}

for s_name, s_open_date in st.session_state.stores.items():
    for task in st.session_state.master_tasks:
        task_date = s_open_date + datetime.timedelta(days=task["offset"])
        date_str = task_date.strftime("%Y-%m-%d")
        offset_val = task['offset']
        d_day_str = f"D{offset_val}" if offset_val < 0 else ("D-Day" if offset_val == 0 else f"D+{offset_val}")
        
        task_id = f"{s_name}_{task['부서']}_{task['주요업무']}"
        default_send_date = task_date - datetime.timedelta(days=1)
        if task_id not in st.session_state.mail_schedules:
            st.session_state.mail_schedules[task_id] = default_send_date

        all_schedule_data.append({
            "task_id": task_id,
            "일자": date_str, "D-Day": d_day_str, "매장명": s_name,
            "부서": task["부서"], "주요업무": task["주요업무"], "담당자": task["담당"],
            "raw_date": task_date, "year": task_date.year, "month": task_date.month
        })
        if date_str not in schedule_map:
            schedule_map[date_str] = []
        schedule_map[date_str].append({"store": s_name, "dept": task["부서"], "task": task["주요업무"]})

# 4. 메인 화면 탭
tab1, tab2, tab3, tab4 = st.tabs(["🗓️ 월별 달력", "📋 전체 공정표", "📮 자동 예약 발송 관리", "👤 이메일 주소록"])

# --- TAB 1: 월별 달력 ---
with tab1:
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

    options = ["🗓️ PC용 넓은 달력 보기", "📱 모바일용 카드 보기"]
    current_index = options.index(st.session_state.view_mode_choice)
    
    selected_mode = st.radio("🖥️ 화면 보기 방식 선택", options, index=current_index, horizontal=True)
    st.session_state.view_mode_choice = selected_mode

    if st.session_state.view_mode_choice == "🗓️ PC용 넓은 달력 보기":
        cal = calendar.Calendar(firstweekday=0)
        month_days = cal.monthdatescalendar(v_year, v_month)
        cols_header = st.columns(7)
        days_name = ["월", "화", "수", "목", "금", "토", "일"]
        for i, d_name in enumerate(days_name):
            cols_header[i].markdown(f"**<center>{d_name}</center>**", unsafe_allow_html=True)
        
        for week in month_days:
            cols = st.columns(7)
            for i, d in enumerate(week):
                d_str = d.strftime("%Y-%m-%d")
                is_cur_month = (d.month == v_month)
                box_style = "border:1px solid #ddd; padding:5px; min-height:110px; border-radius:5px;"
                box_style += "background-color:#ffffff;" if is_cur_month else "background-color:#f9f9f9; color:#bbb;"
                
                tasks_html = ""
                if d_str in schedule_map:
                    for item in schedule_map[d_str]:
                        c_info = get_dept_color(item["dept"])
                        tasks_html += f"""
                        <div style='margin-top:4px; padding:4px; border-radius:4px; background-color:{c_info["bg"]}; border:1px solid {c_info["border"]};'>
                            <div style='font-size:11px; font-weight:bold; color:{c_info["text"]}; border-bottom:1px solid {c_info["border"]}; padding-bottom:1px; margin-bottom:2px;'>
                                [{item["store"]}] {item["dept"]}
                            </div>
                            <div style='font-size:11px; color:#333333; line-height:1.2;'>{item["task"]}</div>
                        </div>"""
                cols[i].markdown(f"<div style='{box_style}'><b>{d.day}</b>{tasks_html}</div>", unsafe_allow_html=True)

    else:
        month_tasks = [s for s in all_schedule_data if s["year"] == v_year and s["month"] == v_month]
        if month_tasks:
            sorted_m_tasks = sorted(month_tasks, key=lambda x: x["일자"])
            for task_item in sorted_m_tasks:
                c_info = get_dept_color(task_item["부서"])
                st.markdown(f"""
                <div class="mobile-card" style="border-left-color: {c_info['border']};">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 5px;">
                        <span style="font-size: 15px; font-weight: bold; color: #1F4E78;">📅 {task_item['일자']}</span>
                        <span class="badge" style="background-color: {c_info['bg']}; color: {c_info['text']}; border: 1px solid {c_info['border']};">
                            [{task_item['매장명']}] {task_item['부서']} ({task_item['D-Day']})
                        </span>
                    </div>
                    <div style="font-size: 13px; color: #333; font-weight: 500;">
                        {task_item['주요업무']}
                    </div>
                    <div style="font-size: 11px; color: #777; margin-top: 4px;">
                        담당: {task_item['담당자']}
                    </div>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.info("해당 월에 예정된 공정이 없습니다.")

# --- TAB 2: 전체 공정표 ---
with tab2:
    st.subheader("📋 전체 지점 오픈 일정 목록")
    if all_schedule_data:
        sorted_schedule = sorted(all_schedule_data, key=lambda x: x["일자"])
        
        html_table = "<table style='width:100%; border-collapse:collapse; border:1px solid #ddd; font-size:14px;'>"
        html_table += "<tr style='background-color:#1F4E78; color:white; font-weight:bold; text-align:center; height:35px;'>"
        html_table += "<th>일자</th><th>D-Day</th><th>매장명</th><th>부서</th><th>주요업무</th><th>담당자</th></tr>"
        for row in sorted_schedule:
            c_info = get_dept_color(row['부서'])
            html_table += f"<tr style='border-bottom:1px solid #ddd; height:32px; text-align:center;'>"
            html_table += f"<td><b>{row['일자']}</b></td><td><span style='background:#f2f2f2; padding:2px 6px; border-radius:3px;'>{row['D-Day']}</span></td>"
            html_table += f"<td><b>{row['매장명']}</b></td>"
            html_table += f"<td><span style='background-color:{c_info['bg']}; color:{c_info['text']}; padding:2px 6px; border-radius:3px; font-weight:bold;'>{row['부서']}</span></td>"
            html_table += f"<td style='text-align:left;'>{row['주요업무']}</td><td>{row['담당자']}</td></tr>"
        html_table += "</table>"
        st.markdown(html_table, unsafe_allow_html=True)

# --- TAB 3: 자동 예약 발송 관리 ---
with tab3:
    st.subheader("📮 공정별 알림 이메일 예약 발송 관리")
    col_g1, col_g2 = st.columns(2)
    user_gmail = col_g1.text_input("발신용 Gmail 주소", value="your_email@gmail.com")
    user_app_pass = col_g2.text_input("Gmail 앱 비밀번호 (16자리)", type="password", value="")

    st.markdown("---")
    if all_schedule_data:
        sorted_schedule = sorted(all_schedule_data, key=lambda x: x["일자"])
        for item in sorted_schedule:
            t_id = item["task_id"]
            current_send_date = st.session_state.mail_schedules.get(t_id, item["raw_date"] - datetime.timedelta(days=1))
            dept_contact = st.session_state.contacts.get(item["담당자"], {"name": "미정", "email": "미등록"})
            
            with st.expander(f"📌 [{item['매장명']}] [{item['부서']}] {item['주요업무']} ({item['일자']})"):
                st.write(f"**담당 팀**: {item['담당자']} ({dept_contact['name']}) | `{dept_contact['email']}`")
                new_send_date = st.date_input("📧 메일 예약 발송 일자", value=current_send_date, key=f"send_date_{t_id}")
                st.session_state.mail_schedules[t_id] = new_send_date
                
                if st.button("🚀 지금 즉시 테스트 발송", key=f"test_send_{t_id}"):
                    if not user_gmail or not user_app_pass:
                        st.error("상단에 Gmail 주소와 앱 비밀번호를 입력하세요.")
                    elif dept_contact['email'] == "미등록":
                        st.error("담당자 메일 주소가 미등록 상태입니다.")
                    else:
                        subject = f"🔔 [알림] [{item['매장명']}] {item['부서']} 공정 안내 ({item['일자']})"
                        body = f"안녕하세요 {dept_contact['name']} 님,\n\n[{item['매장명']}] 지점의 {item['부서']} 공정({item['주요업무']}) 예정일이 {item['일자']} ({item['D-Day']}) 로 예정되어 있습니다."
                        success, msg = send_email_notification(user_gmail, user_app_pass, dept_contact['email'], subject, body)
                        if success: st.success("발송 성공!")
                        else: st.error(msg)

# --- TAB 4: 이메일 주소록 ---
with tab4:
    st.subheader("👤 그룹사 담당자 이메일 주소록 관리")
    c_add1, c_add2, c_add3 = st.columns(3)
    team_input = c_add1.text_input("팀명 (예: 인테리어팀)")
    name_input = c_add2.text_input("담당자 성함/직급", value="홍길동 팀장")
    email_input = c_add3.text_input("이메일 주소", value="hong@9block.co.kr")

    if st.button("➕ 담당자 등록 / 수정"):
        if team_input.strip() and email_input.strip():
            st.session_state.contacts[team_input.strip()] = {
                "name": name_input.strip(),
                "email": email_input.strip()
            }
            save_contacts_to_gsheets()
            st.success("등록 완료!")
            st.rerun()

    st.markdown("---")
    if st.session_state.contacts:
        for t_name, info in list(st.session_state.contacts.items()):
            st.markdown(f"**{t_name}** | {info['name']} | `{info['email']}`")
