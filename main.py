from openpyxl import load_workbook
import os
import shutil


def get_list_of_files(cwd):
    files = os.listdir(cwd)
    xls_list = filter(lambda x: x.endswith('.xlsx') or x.endswith('.xls'), files)
    return xls_list


def check_first_sheet(sheets):
    if sheets[0] == 'Содержание':
        return True
    return False


def delete_string(l_of_files, cwd):
    for f in l_of_files:
        cur_book = load_workbook(f'{cwd}\\{f}')
        print(f'####### Обработка {f} ########')
        sheets = cur_book.sheetnames
        skip = check_first_sheet(sheets=sheets)
        for sh in cur_book:
            if skip:
                skip = False
                continue
            sh.delete_rows(1)
            print(f'Удалена строка в файле {f} на листе "{sh.title}".')
        cur_book.save(f'{cwd}\\Output\\{f}')


def create_output_dir(cwd):
    if os.path.exists(f'{cwd}\\Output'):
        shutil.rmtree(f'{cwd}\\Output')
        os.mkdir(f'{cwd}\\Output')
    else:
        os.mkdir(f'{cwd}\\Output')


def work_with_files():
    cwd = f'{os.getcwd()}\\Input'
    list_of_files = get_list_of_files(cwd)
    create_output_dir(cwd)
    delete_string(list_of_files, cwd)


if __name__ == '__main__':
    work_with_files()
