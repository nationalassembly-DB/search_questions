"""
main 함수. 주질의를 추출합니다.
"""

import os

from module.create_excel import write_excel, load_excel


def main():
    """주질의를 추출합니다"""
    print("\n>>>>>>주질의 추출기<<<<<<\n")
    print("-"*24)
    input_path = input("PDF파일이 존재하는 폴더 경로를 입력하세요 : (종료는 0을 입력)").strip()

    if input_path == '0':
        return

    output_path = input("저장할 엑셀 파일 경로를 입력하세요 : ").strip()

    if not os.path.isdir(input_path):
        print("입력 폴더의 경로를 다시 한번 확인하세요")
        return main()
    if not str(output_path).endswith('.xlsx'):
        output_path = output_path + '.xlsx'
    write_excel(load_excel(output_path), input_path, output_path)
    print("\n~~~모든 작업이 완료되었습니다~~~")

    return main()


if __name__ == "__main__":
    main()
