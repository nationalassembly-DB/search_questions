"""
main 함수. 주질의를 추출합니다.
"""

import os

from module.create_excel import write_excel, load_excel


def main():
    """주질의를 추출합니다"""
    print("-"*24)
    print("\n>>>>>>주질의 추출기<<<<<<\n")
    print("-"*24)
    input_path = input("PDF파일이 존재하는 폴더 경로를 입력하세요(종료는 0을 입력) : ").strip()

    if input_path == '0':
        return 0

    output_path = input(
        "엑셀파일 경로를 입력하세요(확장자포함. 파일이 존재하지 않을 경우 새로 생성) : ").strip()

    if not os.path.isdir(input_path):
        print("입력 폴더의 경로를 다시 한번 확인하세요")
        return main()

    bookmark_level = input("추출할 북마크 LEVEL을 입력하세요 : ")

    if not bookmark_level.isdecimal() or int(bookmark_level) <= 0:
        print("숫자를 입력해주세요")
        return main()

    write_excel(load_excel(output_path), input_path,
                output_path, int(bookmark_level))
    print(f"{output_path}에 주질의 목록이 생성되었습니다.")
    print("\n~~~모든 작업이 완료되었습니다~~~")

    return main()


if __name__ == "__main__":
    main()
