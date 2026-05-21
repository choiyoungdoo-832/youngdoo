import win32com.client
import datetime
import os

# 파일을 임시로 저장할 폴더 설정
DOWNLOAD_DIR = os.path.join(os.getcwd(), "Downloaded_PPTs")
if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

# 💡 병합 순서 및 제출 여부 확인을 위한 기준 데이터
EXPECTED_SENDERS = {
    "허용민": {"order": 1, "team": "연구기획팀"},
    "김태원": {"order": 2, "team": "심혈관팀"},
    "한예지": {"order": 3, "team": "급성감염팀"},
    "정진용": {"order": 4, "team": "Cancer팀"},
    "이소희": {"order": 5, "team": "호르몬팀"},
    "김영은": {"order": 6, "team": "치료용항체팀"},
    "김세희": {"order": 7, "team": "갑상선팀"},
    "함은선": {"order": 8, "team": "당뇨팀"}
}


def get_sort_key(sender_name):
    """보낸 사람 이름을 확인하여 병합 순서 번호를 반환합니다."""
    for name, info in EXPECTED_SENDERS.items():
        if name in sender_name:
            return info["order"]
    return 99  # 명단에 없는 직원은 마지막 순서로 배치


def download_todays_date_ppts():
    """오늘 날짜가 파일명에 포함된 PPT를 다운로드합니다."""
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)

    items = inbox.Items
    items.Sort("[ReceivedTime]", True)

    today = datetime.date.today()
    format_yymmdd = today.strftime("%y%m%d")  # "260521"
    format_yyyymmdd = today.strftime("%Y%m%d")  # "20260521"
    typo_format = "2060521"

    saved_ppt_info = []

    for item in items:
        if item.Class != 43:
            continue

        mail_date = item.ReceivedTime.date()
        if mail_date < today:
            break
        elif mail_date > today:
            continue

        for attachment in item.Attachments:
            filename = attachment.FileName

            if filename.lower().endswith(('.ppt', '.pptx')):
                if (format_yymmdd in filename) or (format_yyyymmdd in filename) or (typo_format in filename):
                    time_str = item.ReceivedTime.strftime("%H%M%S")
                    safe_filename = f"{time_str}_{filename}"
                    save_path = os.path.abspath(os.path.join(DOWNLOAD_DIR, safe_filename))

                    attachment.SaveAsFile(save_path)
                    saved_ppt_info.append({
                        "path": save_path,
                        "sender": item.SenderName,
                        "filename": safe_filename
                    })
                    print(f"[다운로드 완료] 보낸사람: {item.SenderName} -> 파일: {safe_filename}")

    return saved_ppt_info


def check_missing_reports(ppt_info_list):
    """다운로드된 내역을 바탕으로 미제출 팀을 파악하여 안내합니다."""
    # 1. 제출한 사람 목록 추출
    submitted_names = set()
    for info in ppt_info_list:
        sender_name = info['sender']
        for expected_name in EXPECTED_SENDERS.keys():
            if expected_name in sender_name:  # 아웃룩 이름에 '허용민' 등이 포함되어 있는지 확인
                submitted_names.add(expected_name)

    # 2. 미제출자 분류
    missing_list = []
    for name, info in EXPECTED_SENDERS.items():
        if name not in submitted_names:
            missing_list.append((info["team"], name))

    # 3. 제출 현황 출력
    print("\n" + "=" * 50)
    print("📊 주간보고서 제출 현황 요약")
    print("=" * 50)
    print(f"✅ 총 제출: {len(submitted_names)}팀")
    print(f"❌ 미제출: {len(missing_list)}팀")

    if missing_list:
        print("\n[미제출 부서 및 담당자 목록]")
        for team, name in missing_list:
            print(f" 🚨 {team} ({name})")
        print("\n💡 알림: 위 부서의 담당자에게 주간보고서 제출 확인 및 독촉 메일 송부가 필요합니다.")
    else:
        print("\n🎉 모든 팀(8팀)이 주간보고서 제출을 완료했습니다!")
    print("=" * 50 + "\n")


def merge_presentations(ppt_info_list, output_filename="주간보고병합.pptx"):
    """다운로드된 PPT 파일들을 지정된 순서대로 병합합니다."""
    if not ppt_info_list:
        print("병합할 파일이 없어 작업을 종료합니다.")
        return

    sorted_ppt_info = sorted(ppt_info_list, key=lambda x: get_sort_key(x['sender']))
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")

    try:
        merged_presentation = ppt_app.Presentations.Add()
        for info in sorted_ppt_info:
            abs_path = os.path.abspath(info['path'])
            insert_index = merged_presentation.Slides.Count
            merged_presentation.Slides.InsertFromFile(abs_path, insert_index)
            print(f"  -> 슬라이드 복사 완료: {info['filename']}")

        output_path = os.path.abspath(os.path.join(os.getcwd(), output_filename))
        merged_presentation.SaveAs(output_path)
        print(f"\n✅ 최종 병합 완료! 저장 경로: {output_path}")

    except Exception as e:
        print(f"\n❌ 병합 중 오류가 발생했습니다: {e}")
    finally:
        merged_presentation.Close()
        ppt_app.Quit()


# ==========================================
# 실행부
# ==========================================
if __name__ == "__main__":
    print(f"[{datetime.date.today()} 기준] 메일 검색 및 다운로드를 시작합니다...\n")

    # 1. PPT 다운로드
    downloaded_info = download_todays_date_ppts()

    # 2. 💡 [신규] 제출 현황 확인 및 미제출자 안내
    check_missing_reports(downloaded_info)

    # 3. PPT 병합
    merge_presentations(downloaded_info, output_filename="주간보고병합.pptx")
