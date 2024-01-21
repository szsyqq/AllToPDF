# -----------------------------------------------------------------------------------------------------------------
# AllToPDF
# 将Excel文件中的地址转换为地理编码信息

# 详细介绍：


# 使用方法：


# 注意：


# © Yuping Pan 2023-2024
# -----------------------------------------------------------------------------------------------------------------

import os
from natsort import natsorted
import fitz  # 导入本模块需安装pymupdf库
from PIL import Image
from docx2pdf import convert
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# Color print
from termcolor import colored, cprint


def get_file_name_listdir(file_dir):
    return os.listdir(file_dir)  # 不仅仅是文件，当前目录下的文件夹也会被认为遍历到


def is_folder2(enter_path, enter_file_name):
    path = enter_path + "/" + enter_file_name
    if os.path.isdir(path):
        return True
    else:
        return False

def get_file_extension(file_path):
    """
    This function returns the file extension for a given file.

    Returns:
    str: The file extension.
    """
    return os.path.splitext(file_path)[1][1:]

def get_file_name(file_path):
    """
    This function returns the file extension for a given file.

    Returns:
    str: The file extension.
    """
    return os.path.splitext(file_path)[0]

# 处理Word文件 DOC，DOCX


# 处理TXT文件
def create_word_file(txt_file,_temp_path):
    doc = Document()

    template = Document('template.docx')
    for paragraph in template.paragraphs:
        if "{content}" in paragraph.text:  # 替换文本
            new_paragraph = doc.add_paragraph()
            new_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

            full_content = txt_file

            run = new_paragraph.add_run(paragraph.text.replace("{content}", full_content))
            run.font.size = Pt(22)
            run.font.name = 'Times New Romans'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            run.font.bold = False

            # Insert empty paragraphs before the new paragraph
            for _ in range(5):
                empty_paragraph = new_paragraph.insert_paragraph_before("")

    doc.save(_temp_path)

# # 针对二级目录
def folder_in_pdf_out(path):
    in_folder = get_file_name_listdir(path)
    in_folder = natsorted(in_folder)

    _forbid_list = [".DS_Store", "Thumbs.db"]

    for __item in _forbid_list:
        try:
            in_folder.remove(__item)
        except ValueError:
            pass

    doc = fitz.open()

    for in_file in in_folder:

        in_file_path = path + '/' + in_file

        if is_folder2(path, in_file):
            # 如果里面还是文件夹，则内归继续
            initial_pdf = folder_in_pdf_out(in_file_path)
            doc.insertPDF(initial_pdf)
            # print("Success:" + in_file)

        else:
            in_file_kind = get_file_extension(path + '/' + in_file)
            in_file_name = get_file_name(in_file)
            # print('in_file_kind is:',in_file_kind)

            try:
                if in_file_kind == "jpg" \
                        or in_file_kind == "JPEG"  or in_file_kind == "jpeg" \
                        or in_file_kind == "png" or in_file_kind == "gif":
                    try:
                        im = Image.open(in_file_path)
                        im = im.rotate(0, expand=1)
                        try:
                            im.save(in_file_path)
                        except OSError:
                            im = im.convert('RGB')
                            im.save(in_file_path)

                        imgdoc = fitz.open(in_file_path)
                        pdfbytes = imgdoc.convert_to_pdf()
                        imgpdf = fitz.open("pdf", pdfbytes)
                        doc.insert_pdf(imgpdf)

                        print("Success:" + in_file)

                    except:
                        _print_text = "ERR Loading:" + in_file
                        print(colored(_print_text, 'red'))

                elif in_file_kind == "pdf" or in_file_kind == "PDF":
                    try:
                        stream = bytearray(open(in_file_path, "rb").read())
                        pdf_fitz = fitz.open("pdf", stream)
                        doc.insert_pdf(pdf_fitz)
                        print("Success:" + in_file)
                    except:
                        _print_text = "ERR Loading:" + in_file
                        print(colored(_print_text, 'red'))

                elif in_file_kind == "doc" or in_file_kind == "docx" :
                    try:
                        temp_path = '2 Temp' + '/' + in_file +'.pdf'
                        convert(in_file_path, temp_path)

                        stream = bytearray(open(temp_path, "rb").read())
                        pdf_fitz = fitz.open("pdf", stream)
                        doc.insert_pdf(pdf_fitz)

                        print("Success:" + in_file)
                    except:
                        _print_text = "ERR Loading:" + in_file
                        print(colored(_print_text, 'red'))

                elif in_file_kind == "txt" :
                    try:
                        temp_path_txt_to_word = '2 Temp' + '/' + in_file_name +'.docx'
                        create_word_file(in_file_name,temp_path_txt_to_word)

                        temp_path_word_to_pdf = '2 Temp' + '/' + in_file_name + '.pdf'
                        convert(temp_path_txt_to_word,temp_path_word_to_pdf)

                        stream = bytearray(open(temp_path_word_to_pdf, "rb").read())
                        pdf_fitz = fitz.open("pdf", stream)
                        doc.insert_pdf(pdf_fitz)

                        print("Success:" + in_file)
                    except:
                        _print_text = "ERR Loading:" + in_file
                        print(colored(_print_text, 'red'))

                else:
                    _print_text ="ERR Unknown Type:" + in_file
                    print(colored(_print_text, 'red'))

            except AttributeError:
                _print_text = "ERR: AttributeError" + in_file
                print(colored(_print_text, 'red'))

    return doc

# 参数设置 ---------------------------------------------------------------------------

file_path = "/Users/panyp/PycharmProjects/AllToPDF/1 Folders"
output_pdf_path = "/Users/panyp/PycharmProjects/AllToPDF/3 Output"


# 主程序 ------------------------------------------------------------------------------

folder = get_file_name_listdir(file_path)
folder = natsorted(folder)

error_count = 0

_forbid_list = [".DS_Store", "Thumbs.db"]

for __item in _forbid_list:
    try:
        folder.remove(__item)
    except ValueError:
        pass

for _folder in folder:
    folder_file_path = file_path + '/' + _folder

    pdf = folder_in_pdf_out(folder_file_path)

    # 保存PDF
    try:
        pdf.save(output_pdf_path + '/' + _folder + '.pdf')
    except ValueError:
        print(colored(f"ERR:{_folder}文件夹为空",'yellow'))
        error_count = error_count + 1


print(colored(f"错误个数:{error_count}",'yellow'))