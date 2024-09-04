"""
스크립트의 이상점이 생길 경우를 대비하여 log를 생성합니다.
"""

import os


def logging(e, input_path, output_path):
    """이상이 발생한 부분을 log로 저장합니다"""
    output_path = os.path.dirname(output_path)
    log_file_path = os.path.join('\\\\?\\', output_path, 'log.txt')
    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        log_file.write(f'{e} {input_path} {output_path}\n')
