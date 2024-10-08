import os
import comtypes.client
import shutil  # 导入 shutil 库用于文件复制
import time
from random import randint

def batch_convert_docs_to_txt(doc_paths, failed_files, failed_target_folder):
    """
    批量处理多个 .doc 文件，将其内容直接写入 .txt 文件中，并将未成功转换的文件复制到目标文件夹。
    """
    # 创建 Word 应用程序对象
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0  # 禁用警告和弹窗

    for doc_path in doc_paths:
        txt_path = os.path.splitext(doc_path)[0] + '.txt'
        if os.path.exists(txt_path):
            print(f"Already converted: {doc_path}")
            continue

        try:
            print(f"Processing: {doc_path}")
            # 打开 .doc 文件
            doc = word.Documents.Open(doc_path)

            # 提取文档内容并写入到 .txt 文件
            with open(txt_path, 'w', encoding='utf-8') as txt_file:
                for paragraph in doc.Paragraphs:
                    txt_file.write(paragraph.Range.Text.strip() + '\n')

            print(f"Converted {doc_path} to {txt_path}")

        except Exception as e:
            print(f"Failed to convert {doc_path}: {str(e)}")
            failed_files.append(doc_path)

            # 将转换失败的文件复制到目标文件夹
            if not os.path.exists(failed_target_folder):
                os.makedirs(failed_target_folder)  # 如果目标文件夹不存在，则创建文件夹
            shutil.copy(doc_path, failed_target_folder)  # 复制 .doc 文件到目标文件夹
            print(f"Copied failed file {doc_path} to {failed_target_folder}")
            
        finally:
            try:
                if doc:
                    doc.Close(False)  # False 表示不保存更改
            except Exception as e_close:
                print(f"Failed to close {doc_path}: {str(e_close)}")
                failed_files.append(doc_path)  # 记录失败的文件
                word = restart_word(word)  # 重启 Word 应用程序

    # 确保 Word 应用程序被正确退出
    try:
        word.Quit()
        time.sleep(0.5)  # 等待 Word 完全退出
    except Exception as e_quit:
        print(f"Failed to quit Word application: {str(e_quit)}")

def restart_word(word):
    """
    重启 Word 应用程序，释放当前对象并重新创建。
    """
    try:
        word.Quit()
        time.sleep(0.5)  # 等待 Word 完全退出
    except Exception as e_quit:
        print(f"Failed to quit Word application: {str(e_quit)}")

    # 重新创建 Word 应用程序对象
    new_word = comtypes.client.CreateObject('Word.Application')
    new_word.Visible = False
    new_word.DisplayAlerts = 0
    return new_word

def process_all_docs(root_folder):
    """
    遍历指定的根文件夹，将所有 .doc 文件批量处理，并将其内容写入 .txt 文件中。
    """
    # 收集所有 .doc 文件的路径
    doc_files = [os.path.join(root, file)
                 for root, dirs, files in os.walk(root_folder)
                 for file in files if file.endswith('.doc') and not file.endswith('.docx')]

    # 存储转换失败的文件列表
    failed_files = []

    a = randint(0,100)

    # 定义存放未成功转换文件的目标文件夹路径
    failed_target_folder = os.path.join(root_folder, 'FailedFiles'+str(a))

    # 分批次处理文件（比如每次处理 3 个文件）
    batch_size = 4
    for i in range(0, len(doc_files), batch_size):
        batch_convert_docs_to_txt(doc_files[i:i + batch_size], failed_files, failed_target_folder)

# 设置根文件夹路径（根据实际情况修改）
root_folder = r'C:\Users\hansh\Desktop\lyy\招商引资政策文本库\招商引资政策文本库\地方政策文本库-按时间顺序（3415篇）'

# 执行批量转换
process_all_docs(root_folder)
