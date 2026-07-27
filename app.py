import datetime
import smtplib
import json
import os
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 환경변수 또는 Secrets에서 정보 가져오기
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID")
GCP_CREDS_JSON = os.environ.get("GCP_SERVICE_ACCOUNT")
SENDER_EMAIL = os.environ.get("SENDER_EMAIL")
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD")

def run_auto_mailer():
    if not all([SPREADSHEET_ID, GCP_CREDS_JSON, SENDER_EMAIL, SENDER_PASSWORD]):
        print("❌ 필요한 환경변수(Secrets)가 설정되지 않았습니다.")
        return

    sender_email = SENDER_EMAIL.strip()
    sender_password = SENDER_PASSWORD.strip().replace(" ", "") # 앱 비밀번호 공백 제거

    try:
        # 1. 구글 시트 연결
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds_dict = json.loads(GCP_CREDS_JSON)
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        
        doc = client.open_by_key(SPREADSHEET_ID)

        # 2. 지점 정보 불러오기
        ws_stores = doc.worksheet("stores")
        stores_data = ws_stores.get_all_records()
        stores = {}
        for r in stores_data:
            s_name = str(r.get("매장명", "")).strip()
            s_date = str(r.get("오픈일", "")).strip()
            s_type = str(r.get("구분", "가맹")).strip()
            if s_name and s_date:
                stores[s_name] = {
                    "date": datetime.datetime.strptime(s_date, "%Y-%m-%d").date(),
                    "type": s_type
                }

        # 3. 주소록 불러오기 (전체 임직원의 개별 이메일 주소 리스트 추출)
        ws_contacts = doc.worksheet("contacts")
        contacts_data = ws_contacts.get_all_records()
        
        all_individual_emails = []
        email_regex = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
        
        for r in contacts_data:
            email = str(r.get("이메일", "")).strip()
            if email and re.match(email_regex, email):
                if email not in all_individual_emails:
                    all_individual_emails.append(email)

        if not all_individual_emails:
            print("⚠️ 주소록 시트에 등록된 유효한 이메일 주소가 없습니다.")
            return

        print(f"📧 등록된 전체 개별 이메일 리스트 ({len(all_individual_emails)}명): {all_individual_emails}")

        # 4. 작업 수정/완료 상태(task_status) 불러오기
        task_status = {}
        try:
            ws_status = doc.worksheet("task_status")
            status_data = ws_status.get_all_records()
            for r in status_data:
                t_id = str(r.get("task_id", "")).strip()
                if t_id:
                    task_status[t_id] = {
                        "completed": bool(r.get("completed", False)),
                        "custom_task": str(r.get("custom_task", "")),
                        "custom_offset": int(r.get("custom_offset", 0)) if str(r.get("custom_offset", "")).strip() != "" else None,
                        "custom_assignee": str(r.get("custom_assignee", ""))
                    }
        except Exception:
            print("💡 'task_status' 시트가 비어있거나 없습니다.")

        # 5. 마스터 공정 목록
        master_tasks = [
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

        today = datetime.date.today()
        print(f"📅 오늘 날짜 기준({today}) 자동 발송 대상 확인 시작...")

        # 6. D-1일 항목 검사 및 등록된 모든 개별 이메일로 수신 지정하여 발송
        for s_name, s_info in stores.items():
            s_open_date = s_info["date"]
            s_type = s_info["type"]
            
            for task in master_tasks:
                task_id = f"{s_name}_{task['부서']}_{task['주요업무']}"
                status_info = task_status.get(task_id, {})
                
                # 이미 완료 처리된 공정은 자동 알림 제외
                if status_info.get("completed", False):
                    continue

                # 커스텀 변경 사항 반영
                task_title = status_info.get("custom_task") if status_info.get("custom_task") else task["주요업무"]
                offset_val = status_info.get("custom_offset") if status_info.get("custom_offset") is not None else task["offset"]
                assignee_val = status_info.get("custom_assignee") if status_info.get("custom_assignee") else task["담당"]

                task_date = s_open_date + datetime.timedelta(days=offset_val)
                send_target_date = task_date - datetime.timedelta(days=1)  # D-1일 발송
                
                # 오늘이 D-1 알림일인 경우
                if send_target_date == today:
                    subject = f"🔔 [공정 예정 알림] [{s_type}|{s_name}] {task['부서']} - {task_title}"
                    body = f"안녕하세요 팀원 여러분,\n\n[{s_name}] ({s_type}) 지점의 내일 예정 공정 안내입니다.\n\n"
                    body += f"📌 지점명: {s_name} ({s_type})\n"
                    body += f"📌 공정명: {task_title}\n"
                    body += f"📌 주관부서/담당: {assignee_val}\n"
                    body += f"📌 공정예정일: 내일 ({task_date})\n\n"
                    body += f"해당 부서 및 관계자분들께서는 일정을 확인하시어 오픈 준비에 차질이 없도록 진행해 주시기 바랍니다."
                    
                    try:
                        msg = MIMEMultipart()
                        msg['From'] = sender_email
                        # 등록된 모든 임직원 개별 이메일을 콤마(,)로 연결하여 수신자로 지정
                        msg['To'] = ", ".join(all_individual_emails)
                        msg['Subject'] = subject
                        msg.attach(MIMEText(body, 'plain', 'utf-8'))

                        server = smtplib.SMTP("smtp.gmail.com", 587, timeout=20)
                        server.ehlo()
                        server.starttls()
                        server.login(sender_email, sender_password)
                        server.sendmail(sender_email, all_individual_emails, msg.as_string())
                        server.quit()
                        
                        print(f"✅ 전체 개별 인원 알림 발송 성공 ➔ [{s_name}] {task_title} (수신자 {len(all_individual_emails)}명)")
                    except Exception as e:
                        print(f"❌ 메일 발송 실패: {e}")

    except Exception as e:
        print(f"❌ 전체 시스템 오류: {e}")

if __name__ == "__main__":
    run_auto_mailer()
