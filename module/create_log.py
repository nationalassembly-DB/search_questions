import os


def logging(e, input_path, output_path):
    output_path = os.path.dirname(output_path)
    log_file_path = os.path.join('\\\\?\\', output_path, 'log.txt')
    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        log_file.write(f'{e} {input_path} {output_path}\n')
