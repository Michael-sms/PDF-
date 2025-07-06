from docx2pdf import convert
import argparse

def get_pdf(input_file):
    #获取文件名
    file_name = input_file.split(".docx")[0]
    #转换为pdf
    convert(input_file, file_name+".pdf")
    #返回pdf文件路径
    return file_name+".pdf"

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Convert Word document to PDF.')
    parser.add_argument('input_file', type=str, help='docx文件路径')
    args = parser.parse_args()
    pdf_file = get_pdf(args.input_file)

    print("文件成功转换为pdf，保存在：", pdf_file)

