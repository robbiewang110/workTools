import os

from PIL import Image
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


def convert_images_to_pdf(images_folder, output_pdf_path):
    index = 1
    filelist = []

    try:
        os.mkdir("./tmp/")
    except (IOError, OSError) as e:
        # 记录错误日志
        print(f"错误信息: {str(e)}")

    # 遍历指定文件夹中的所有图片文件
    for filename in os.listdir(images_folder):
        if filename.endswith(".jpg") or filename.endswith(".png"):
            image_path = os.path.join(images_folder, filename)

            print("处理文件：" + image_path)
            filename = "./tmp/" + str(index) + ".pdf"
            index = index + 1
            filelist.append(filename)

            try:
                # 打开图片文件
                image = Image.open(image_path)

                # 设置PDF的大小未图片大小
                pdfsize = (image.width, image.height)

                # 创建新的PDF页面
                pdf_page = canvas.Canvas(filename, pagesize=pdfsize)
                pdf_page.drawImage(image_path, 0, 0, pdfsize[0], pdfsize[1])
                pdf_page.showPage()
                pdf_page.save()

            except (IOError, OSError) as e:
                # 记录错误日志
                print(f"无法处理图片文件: {image_path}")
                print(f"错误信息: {str(e)}")

    # 将生成的PDF保存到指定路径，然后合并
    try:
        merger = PdfMerger()

        for path in filelist:
            merger.append(path)
        merger.write(output_pdf_path)
        merger.close()

    except (IOError, OSError) as e:
        # 记录错误日志
        print(f"无法保存PDF文件: {output_pdf_path}")
        print(f"错误信息: {str(e)}")
        return False

    # 清理临时文件
    try:
        for path in filelist:
            os.remove(path)
        return True
    except (IOError, OSError) as e:
        # 记录错误日志
        print(f"无法保存PDF文件: {output_pdf_path}")
        print(f"错误信息: {str(e)}")
        return False


def test_convert_images_to_pdf():
    return True


def convert_images_to_pdf2(images_folder, output_pdf_path):
    pdf_writer = PdfWriter()

    # 遍历指定文件夹中的所有图片文件
    for filename in os.listdir(images_folder):
        if filename.endswith(".jpg") or filename.endswith(".png"):
            image_path = os.path.join(images_folder, filename)

            print("处理文件：" + image_path)

            try:
                image = Image.open(image_path)

                img_pos_x = 0
                img_pos_y = 0
                img_width_new = image.height
                img_height_new = image.width

                # 创建新的PDF页面
                pdf_page = canvas.Canvas('temp.pdf', pagesize=A4)

                page_width, page_height = pdf_page._pagesize

                y_ratio = image.height / page_height
                x_ratio = image.width / page_width

                if x_ratio > y_ratio:
                    if x_ratio > 1:
                        img_width_new = page_width
                        img_height_new = image.height / x_ratio
                        img_pos_y = (page_height - img_height_new) / 2
                        img_pos_x = 0
                    else:
                        img_width_new = image.width
                        img_height_new = image.height / x_ratio
                        img_pos_y = (page_height - img_height_new) / 2
                        img_pos_x = 0
                else:
                    print("ss")
                    if y_ratio > 1:
                        img_height_new = page_height
                        img_width_new = image.width / y_ratio
                        img_pos_x = (page_width - img_width_new) / 2
                        img_pos_y = 0
                    else:
                        img_height_new = image.height
                        img_width_new = image.width / y_ratio
                        img_pos_x = (page_width - img_width_new) / 2
                        img_pos_y = 0

                pdf_page.drawImage(image_path, img_pos_x,
                                   img_pos_y, img_width_new, img_height_new)
                pdf_page.showPage()
                pdf_page.save()

                pdf_reader = PdfReader('temp.pdf')
                # 将PDF页面添加到PDF写入器中
                pdf_page_obj = pdf_reader.pages[0]
                pdf_writer.add_page(pdf_page_obj)

                # 删除临时文件
                os.remove('temp.pdf')

            except (IOError, OSError) as e:
                # 记录错误日志
                print(f"无法处理图片文件: {image_path}")
                print(f"错误信息: {str(e)}")
    # 将生成的PDF保存到指定路径
    try:
        with open(output_pdf_path, "wb") as output_pdf:
            pdf_writer.write(output_pdf)
        print(f"成功生成PDF文件: {output_pdf_path}")
        return True

    except (IOError, OSError) as e:
        # 记录错误日志
        print(f"无法保存PDF文件: {output_pdf_path}")
        print(f"错误信息: {str(e)}")
        return False


# 示例用法
images_folder = "./"
output_pdf_path = "result.pdf"

success = convert_images_to_pdf(images_folder, output_pdf_path)

if success:
    print("图片转换为PDF成功！")
else:
    print("图片转换为PDF失败！")
    """_summary_
    """
