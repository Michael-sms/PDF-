from docx2pdf import convert
import argparse
import os

def convert_file(input_file):
    """转换单个文件为PDF"""
    #判断文件是否存在
    if not input_file.lower().endswith('.docx'):
        print("文件格式不正确，请检查后重试")
        return None

    #获取文件名
    file_name = os.path.splitext(input_file)[0]
    output_file = file_name + ".pdf"

    try:
        convert(input_file, output_file)
        print(f"转换成功：{input_file} -> {output_file}")
        return output_file
    except Exception as e:
        print(f"转换失败：{input_file} -> {output_file}，错误信息：{e}")
        return None

def convert_dir(input_dir):
    """转换目录下所有文件"""
    if not os.path.isdir(input_dir):
        print("输入路径不存在或不是目录")
        return []

    output_files = []
    for file_name in os.listdir(input_dir):
        if file_name.lower().endswith('.docx'):
            input_file = os.path.join(input_dir,file_name)
            output_file = convert_file(input_file)
            if output_file:
                output_files.append(output_file)
    print(f"转换完成，共转换{len(output_files)}个.docx文件")

    return output_files


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Convert Word document to PDF.')
    parser.add_argument('input_file', type=str, help='docx文件路径或文件夹路径')
    args = parser.parse_args()

    input_path = args.input_path
    if  os.path.isfile(input_path):
        """转换单个文件"""
        convert_file(input_path)
    elif os.path.isdir(input_path):
        """转换目录下所有文件"""
        convert_dir(input_path)
    else:
        print(f"输入路径{input_path}不存在或不是文件或目录")




