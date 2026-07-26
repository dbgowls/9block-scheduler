import streamlit as st
import datetime
import calendar
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# 1. 페이지 설정 (모바일 최적화 레이아웃)
st.set_page_config(page_title="9BLOCK 스케줄러", page_icon="📱", layout="wide", initial_sidebar_state="collapsed")

# 모바일 커스텀 CSS 적용
st.markdown("""
    <style>
    .stApp { padding: 5px; }
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

st.title("📱 9BLOCK 오픈 스케줄러")

# 초기 표준 템플릿
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

# 초기 그룹사 담당자 주소록
DEFAULT_CONTACTS = {
    "개발팀": {"name": "김개발 팀장", "email": "dev@9block.co.kr"},
    "구매/운영팀": {"name": "이구매 팀장", "email": "buy@9block.co.kr"},
    "인테리어팀": {"name": "박설계 팀장", "email": "interior@9block.co.kr"},
    "마케팅팀": {"name": "최홍보 팀장", "email": "mkt@9block.co.kr"},
    "운영팀": {"name": "정운영 팀장", "email": "ops@9block.co.kr"},
    "전부서": {"name": "오픈지원TF", "email": "tf@9block.co.kr"}
}

# 2. 세션 상태 초기화
if "stores" not in st.session_state:
    st.session_state.stores = {"충청점": datetime.date(2026, 9, 18)}

if "current_view_date" not in st.session_state:
    st.session_state.current_view_date = datetime.date(2026, 8, 1)

if "master_tasks" not in st.session_state:
    st.session_state.master_tasks = [dict(t) for t in DEFAULT_TASKS]

if "contacts" not in st.session_state:
    st.session_state.contacts = dict(DEFAULT_CONTACTS)

if "mail_schedules" not in st.session_state:
    st.session_state.mail_schedules = {}

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
        "개발": {"bg": "#e3f2fd", "text": "#1e88e5", "border": "#1e88e5"},
        "구매/운영": {"bg": "#e8f5e9", "text": "#2e7d32", "border": "#2e7d32"},
        "인테리어": {"bg": "#fff3e0", "text": "#e65100", "border": "#e65100"},
        "마케팅": {"bg": "#f3e5f5", "text": "#8e24aa", "border": "#8e24aa"},
        "전부서": {"bg": "#ffebee", "text": "#c62828", "border": "#c62828"}
    }
    return color_map.get(dept_name, {"bg": "#f5f5f5", "text": "#424242", "border": "#424242"})

# --- 사이드바 메뉴 (모바일 대응) ---
st.sidebar.title("⚙️ 설정 및 편집")
st.sidebar.header("➕ 신규 지점 등록")
with st.sidebar.form("add_store_form", clear_on_submit=True):
    new_store_name = st.text_input("매장명 (예: 강남점)")
    new_open_date = st.date_input("GRAND OPEN 예정일", value=datetime.date(2026, 10, 15))
    if st.form_submit_button("지점 추가하기"):
        if new_store_name.strip():
            st.session_state.stores[new_store_name.strip()] = new_open_date
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
        st.sidebar.success("수정 완료!")
        st.rerun()
    if col_btn2.button("지점 삭제"):
        del st.session_state.stores[selected_edit_store]
        st.sidebar.warning("삭제 완료!")
        st.rerun()

st.sidebar.markdown("---")
st.sidebar.header("⚙️ 공정 항목 추가 / 삭제")
with st.sidebar.expander("➕ 새 공정 항목 추가"):
    with st.form("add_task_form", clear_on_submit=True):
        add_dept = st.selectbox("부서 선택", ["개발", "구매/운영", "인테리어", "마케팅", "전부서"])
        add_title = st.text_input("주요 업무 내용")
        add_offset = st.number_input("D-Day (예: -15)", value=-10, step=1)
        add_team = st.text_input("담당 팀명", value="운영팀")
        if st.form_submit_button("공정 추가"):
            if add_title.strip():
                st.session_state.master_tasks.append({
                    "부서": add_dept, "주요업무": add_title.strip(), "offset": add_offset, "담당": add_team.strip()
                })
                st.rerun()

# 3. 데이터 연산
all_schedule_data = []

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

# 4. 메인 화면 탭 (모바일 최적화 4개 탭)
tab1, tab2, tab3, tab4 = st.tabs(["📱 모바일 달력", "📋 공정표", "📮 알림 예약", "👤 주소록"])

# --- TAB 1: 모바일 최적화 달력 (카드 View / 격자 View 전환) ---
with tab1:
    c_prev, c_title, c_next = st.columns([1, 2, 1])
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
    c_title.markdown(f"<h4 style='text-align: center;'>🗓️ {v_year}년 {v_month:02d}월</h4>", unsafe_allow_html=True)

    # 모바일 전용 부서 필터
    selected_dept_filter = st.radio("부서 필터", ["전체", "개발", "구매/운영", "인테리어", "마케팅", "전부서"], horizontal=True)

    # 선택 월의 일정 필터링
    month_tasks = [
        s for s in all_schedule_data 
        if s["year"] == v_year and s["month"] == v_month and (selected_dept_filter == "전체" or s["부서"] == selected_dept_filter)
    ]
    
    view_mode = st.radio("보기 방식", ["📱 터치 카드 보기", "🗓️ 격자 달력 보기"], horizontal=True)

    if view_mode == "📱 터치 카드 보기":
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
            st.info("해당 월/부서에 예정된 공정이 없습니다.")

    else:
        # 기존 7열 격자 달력
        cal = calendar.Calendar(firstweekday=0)
        month_days = cal.monthdatescalendar(v_year, v_month)
        cols_header = st.columns(7)
        days_name = ["월", "화", "수", "목", "금", "토", "일"]
        for i, d_name in enumerate(days_name):
            cols_header[i].markdown(f"**<center>{d_name}</center>**", unsafe_allow_html=True)
        
        schedule_map = {}
        for item in month_tasks:
            d_str = item["일자"]
            if d_str not in schedule_map: schedule_map[d_str] = []
            schedule_map[d_str].append(item)

        for week in month_days:
            cols = st.columns(7)
            for i, d in enumerate(week):
                d_str = d.strftime("%Y-%m-%d")
                is_cur_month = (d.month == v_month)
                box_style = "border:1px solid #ddd; padding:3px; min-height:80px; border-radius:5px;"
                box_style += "background-color:#ffffff;" if is_cur_month else "background-color:#f9f9f9; color:#bbb;"
                
                tasks_html = ""
                if d_str in schedule_map:
                    for item in schedule_map[d_str]:
                        c_info = get_dept_color(item["부서"])
                        tasks_html += f"""<div style='font-size:10px; background-color:{c_info["bg"]}; color:{c_info["text"]}; margin-top:2px; padding:2px; border-radius:3px;'>[{item['매장명']}] {item['부서']}</div>"""
                cols[i].markdown(f"<div style='{box_style}'><b>{d.day}</b>{tasks_html}</div>", unsafe_allow_html=True)

# --- TAB 2: 공정표 ---
with tab2:
    st.subheader("📋 전체 공정 일정표")
    if all_schedule_data:
        sorted_schedule = sorted(all_schedule_data, key=lambda x: x["일자"])
        for row in sorted_schedule:
            c_info = get_dept_color(row['부서'])
            st.markdown(f"""
            <div style="padding: 8px; border-bottom: 1px solid #eee; font-size: 13px;">
                <b>{row['일자']}</b> <span class="badge" style="background:{c_info['bg']}; color:{c_info['text']};">[{row['매장명']}] {row['부서']}</span> <b>{row['D-Day']}</b><br>
                <span style="color:#444;">{row['주요업무']}</span> <span style="font-size:11px; color:#888;">({row['담당자']})</span>
            </div>
            """, unsafe_allow_html=True)

# --- TAB 3: 메일 예약 발송 ---
with tab3:
    st.subheader("📮 알림 예약 발송")
    user_gmail = st.text_input("발신 Gmail", value="your_email@gmail.com")
    user_app_pass = st.text_input("앱 비밀번호 (16자리)", type="password", value="")

    st.markdown("---")

    if all_schedule_data:
        sorted_schedule = sorted(all_schedule_data, key=lambda x: x["일자"])
        for item in sorted_schedule:
            t_id = item["task_id"]
            current_send_date = st.session_state.mail_schedules.get(t_id, item["raw_date"] - datetime.timedelta(days=1))
            dept_contact = st.session_state.contacts.get(item["담당자"], {"name": "미정", "email": "미등록"})
            
            with st.expander(f"📌 [{item['매장명']}] {item['부서']} - {item['일자']}"):
                st.write(f"**업무**: {item['주요업무']}")
                st.write(f"**수신 메일**: `{dept_contact['email']}`")
                new_send_date = st.date_input("예약 발송일", value=current_send_date, key=f"send_date_{t_id}")
                st.session_state.mail_schedules[t_id] = new_send_date
                
                if st.button("🚀 지금 바로 테스트 발송", key=f"test_send_{t_id}"):
                    if not user_gmail or not user_app_pass:
                        st.error("Gmail 정보를 먼저 입력하세요.")
                    elif dept_contact['email'] == "미등록":
                        st.error("담당자 메일이 등록되지 않았습니다.")
                    else:
                        subject = f"🔔 [알림] [{item['매장명']}] {item['부서']} 공정 안내 ({item['일자']})"
                        body = f"안녕하세요 {dept_contact['name']} 님,\n\n[{item['매장명']}] 지점의 {item['부서']} 공정({item['주요업무']}) 예정일이 {item['일자']} ({item['D-Day']}) 로 예정되어 있습니다."
                        success, msg = send_email_notification(user_gmail, user_app_pass, dept_contact['email'], subject, body)
                        if success: st.success("발송 성공!")
                        else: st.error(msg)

# --- TAB 4: 주소록 ---
with tab4:
    st.subheader("👤 그룹사 담당자 주소록")
    team_input = st.text_input("팀명 (예: 인테리어팀)")
    name_input = st.text_input("담당자 성함", value="홍길동 팀장")
    email_input = st.text_input("이메일 주소", value="hong@9block.co.kr")

    if st.button("➕ 담당자 등록"):
        if team_input.strip() and email_input.strip():
            st.session_state.contacts[team_input.strip()] = {
                "name": name_input.strip(),
                "email": email_input.strip()
            }
            st.success("등록되었습니다.")
            st.rerun()

    st.markdown("---")
    if st.session_state.contacts:
        for t_name, info in list(st.session_state.contacts.items()):
            st.markdown(f"**{t_name}** | {info['name']} | `{info['email']}`")
