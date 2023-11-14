import os

from PIL import Image
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from PIL import Image


def crop_images(source_dir, dest_dir):
    index = 1
    filelist = []

    try:
        os.mkdir(dest_dir)
    except (IOError, OSError) as e:
        # 记录错误日志
        print(f"错误信息: {str(e)}")

    # 遍历指定文件夹中的所有图片文件
    for filename in os.listdir(source_dir):
        if filename.endswith(".jpg") or filename.endswith(".png"):
            image_path = os.path.join(source_dir, filename)

            print("处理文件："+image_path)
            filenameDest = os.path.join(dest_dir, filename)
            index = index+1

            try:
                # 打开图片文件
                image = Image.open(image_path)
                image = image.crop((0, 80, 1920, 1160))
                image.save(filenameDest)

            except (IOError, OSError) as e:
                # 记录错误日志
                print(f"无法处理图片文件: {image_path}")
                print(f"错误信息: {str(e)}")

    return True


# 示例用法
source_dir = "./"
dest_dir = "./NewImages/"

success = crop_images(source_dir, dest_dir)

if success:
    print("图片转换为PDF成功！")
else:
    print("图片转换为PDF失败！")
    """_summary_
    """
