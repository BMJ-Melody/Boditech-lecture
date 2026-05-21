import os
import win32com.client
import datetime


def merge_powerpoint_files(ppt_files, output_path):
    """
    여러 개의 PPT 파일을 하나의 PPT로 병합하는 함수
    """
    if not ppt_files:
        return

    print("\n" + "=" * 60)
    print("🔄 다운로드된 PPT 파일들의 병합을 시작합니다...")

    ppt_app = None
    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        base_ppt = ppt_app.Presentations.Open(ppt_files[0])

        for ppt_file in ppt_files[1:]:
            print(f"  -> [{os.path.basename(ppt_file)}] 슬라이드를 추가하는 중...")
            insert_index = base_ppt.Slides.Count
            base_ppt.Slides.InsertFromFile(ppt_file, insert_index)

        base_ppt.SaveAs(output_path)
        base_ppt.Close()
        print(f"\n✨ 병합 완료! 최종 파일이 생성되었습니다: {output_path}")

    except Exception as e:
        print(f"❌ PPT 병합 중 오류가 발생했습니다: {e}")
    finally:
        if ppt_app:
            ppt_app.Quit()


def get_emails_and_merge_ppts():
    try:
        # 1. 아웃룩 연결 및 받은 편지함 가져오기
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)

        # 2. 이미지 구조에 맞춰 순차적으로 폴더 들어가기
        print("폴더 경로를 탐색합니다: [받은 편지함] -> [2. 영업관리] -> [(나)방민정]")

        target_folder = None
        for sub1 in inbox.Folders:
            if "2. 영업관리" in sub1.Name:
                for sub2 in sub1.Folders:
                    if "(나)방민정" in sub2.Name:
                        target_folder = sub2
                        break
            if target_folder:
                break

        if not target_folder:
            print("❌ '(나)방민정' 폴더를 찾지 못했습니다. 아웃룩이 켜져 있는지 확인해주세요.")
            return

        print(f"✅ 폴더 찾기 성공! [{target_folder.Name}] 폴더에서 메일을 검색합니다.\n")
        print("=" * 60)

        # 3. 자료 취합 대상자 리스트 (순서대로 딕셔너리에 저장)
        expected_list = {
            "허용민": "연구기획팀",
            "김태원": "심혈관팀",
            "한예지": "급성감염팀",
            "정진용": "Cancer팀",
            "이소희": "호르몬팀",
            "김영은": "치료용항체팀",
            "김세희": "갑상선팀",
            "함은선": "당뇨팀"
        }
        # 제출한 사람의 이름을 담을 집합(Set)
        submitted_persons = set()

        # 4. 메일 가져오기 및 최신순 정렬
        items = target_folder.Items
        items.Sort("[ReceivedTime]", True)

        # 오늘 날짜 (YYMMDD)
        today_date = datetime.date.today()
        date_str_yymmdd = today_date.strftime("%y%m%d")

        keywords = ["주간회의자료", "회의자료"]
        count = 0

        # 다운로드 경로 설정
        base_dir = os.getcwd()
        download_folder = os.path.join(base_dir, "다운로드_PPT")
        if not os.path.exists(download_folder):
            os.makedirs(download_folder)

        downloaded_ppt_paths = []

        # 5. 조건 검사 시작
        for msg in items:
            if msg.Class != 43:
                continue

            try:
                msg_date = msg.ReceivedTime.date()
                if msg_date < today_date:
                    break
                if msg_date > today_date:
                    continue

                subject = msg.Subject if msg.Subject else ""
                body = msg.Body if msg.Body else ""

                has_target_ppt = False
                temp_ppt_attachments = []
                keyword_in_filename = False

                # 첨부파일(PPT) 검사
                for att in msg.Attachments:
                    filename = att.FileName
                    if filename.lower().endswith(('.ppt', '.pptx')):
                        if date_str_yymmdd in filename:
                            has_target_ppt = True
                            temp_ppt_attachments.append(att)
                            if any(kw in filename for kw in keywords):
                                keyword_in_filename = True

                if not has_target_ppt:
                    continue

                subject_has_kw = any(kw in subject for kw in keywords)
                body_has_kw = any(kw in body for kw in keywords)

                # 키워드 조건 일치 시 처리
                if subject_has_kw or body_has_kw or keyword_in_filename:
                    sender = msg.SenderName
                    received_time = msg.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")

                    print(f"[{count + 1}] 제목: {subject} (발송자: {sender})")

                    # 💡 제출자 확인 로직: 아웃룩 발송자 이름에 리스트의 담당자 이름이 포함되어 있는지 확인
                    for person_name in expected_list.keys():
                        if person_name in sender:
                            submitted_persons.add(person_name)

                    for att in temp_ppt_attachments:
                        safe_filename = f"{count + 1}_{att.FileName}"
                        save_path = os.path.join(download_folder, safe_filename)
                        att.SaveAsFile(save_path)
                        print(f"    [저장완료] {safe_filename} (오늘 날짜 {date_str_yymmdd} 포함)")
                        downloaded_ppt_paths.append(os.path.abspath(save_path))

                    count += 1

            except Exception as e:
                continue

        if count == 0:
            print(f"오늘 '(나)방민정' 폴더로 수신된 메일 중 조건(날짜: {date_str_yymmdd} 포함 등)에 맞는 메일이 없습니다.")
        else:
            print("-" * 60)
            print(f"메일 검색 및 다운로드 완료! 총 {count}개의 메일에서 PPT를 추출했습니다.")

            # 병합 실행
            if downloaded_ppt_paths:
                final_output_path = os.path.join(base_dir, "주간보고병합.pptx")
                merge_powerpoint_files(downloaded_ppt_paths, final_output_path)

        # 6. 💡 제출 및 미제출 현황 출력
        print("\n" + "=" * 60)
        print("📊 주간 보고서 제출 현황")
        print("-" * 60)

        missing_teams = []
        for person, team in expected_list.items():
            if person not in submitted_persons:
                missing_teams.append(f"{team} ({person})")

        if not missing_teams:
            print("🎉 모든 팀이 주간 보고서를 성공적으로 제출했습니다!")
        else:
            print(f"⚠️ 미제출 팀 (총 {len(missing_teams)}팀):")
            for idx, info in enumerate(missing_teams, 1):
                print(f"  {idx}. {info}")
        print("=" * 60)

    except Exception as e:
        print(f"프로그램 실행 중 오류가 발생했습니다: {e}")


if __name__ == "__main__":
    get_emails_and_merge_ppts()
