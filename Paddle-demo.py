import os
from paddleocr import PaddleOCR
from PIL import Image, ImageDraw, ImageFont
from pdf2image import convert_from_path
from docx import Document
from openpyxl import Workbook

# 定义输入输出路径
input_dir = "/Users/buxiangzun/Desktop/input"  # 输入路径
output_dir = "/Users/buxiangzun/Desktop/output"  # 输出路径

# 输出目录
os.makedirs(output_dir, exist_ok=True)

# 定义ocr
ocr = PaddleOCR(use_angle_cls=True, lang="ch")

# 滑动窗口
slice = {'horizontal_stride': 300, 'vertical_stride': 500, 'merge_x_thres': 50, 'merge_y_thres': 35}

# 编制输入
for filename in os.listdir(input_dir):
    input_path = os.path.join(input_dir, filename)
    file_base, file_ext = os.path.splitext(filename)

    # 创建excel工作簿
    wb = Workbook()
    ws = wb.active
    ws.append(['minx', 'miny', 'maxx', 'maxy', 'confidence', 'text'])  # 表头

    if file_ext.lower() == '.pdf':
        # 若是pdf，转为image
        images = convert_from_path(input_path)
        for page_num, image in enumerate(images):
            img_output_path = os.path.join(output_dir, f"{file_base}_page_{page_num + 1}.jpg")
            image.save(img_output_path, 'JPEG')
            image_path = img_output_path

            # 对转换完的图提取文字
            results = ocr.ocr(image_path, cls=True, slice=slice)

            # 载入原始图片
            original_image = Image.open(image_path).convert("RGB")
            original_width, original_height = original_image.size

            # 创建一个新的画板，两倍原来的图片大小
            new_image = Image.new("RGB", (original_width * 2, original_height), (255, 255, 255))

            # 把原图拼接在新图的左侧
            new_image.paste(original_image, (0, 0))

            # 编辑新图
            draw = ImageDraw.Draw(new_image)

            # 确定字体格式
            font_path = "/System/Library/Fonts/Supplemental/Songti.ttc"
            font = ImageFont.truetype(font_path, size=70)

            # ocr输出结果添加在新图上
            for res in results:
                for line in res:
                    box = [tuple(point) for point in line[0]]
                    min_x = min(point[0] for point in box)
                    min_y = min(point[1] for point in box)
                    max_x = max(point[0] for point in box)
                    max_y = max(point[1] for point in box)
                    txt = line[1][0]
                    score = line[1][1]

                    # 绘制蓝框
                    draw.rectangle([original_width + min_x, min_y, original_width + max_x, max_y], outline="blue", width=2)

                    # 根据蓝框确定字体应该放在哪，目标是中间
                    bbox = draw.textbbox((0, 0), txt, font=font)  # 获取边界
                    text_width = bbox[2] - bbox[0]  # 宽x2-x1
                    text_height = bbox[3] - bbox[1]  # 高y2-y1
                    text_x = original_width + min_x + (max_x - min_x - text_width) // 2
                    text_y = min_y + (max_y - min_y - text_height) // 2

                    # 在框内绘制文字
                    draw.text((text_x, text_y), txt, fill="black", font=font)

                    # 在框底部加置信度
                    draw.text((original_width + min_x, max_y + 5), f" {score:.2f}", fill="green", font=font)

                    # 将数据写入Excel
                    ws.append([min_x, min_y, max_x, max_y, score, txt])

            # 保存最终图片
            annotated_output_path = os.path.join(output_dir, f"{file_base}_page_{page_num + 1}_annotated.jpg")
            new_image.save(annotated_output_path)

            # 保存为xlsx文件
            xlsx_output_path = os.path.join(output_dir, f"{file_base}_page_{page_num + 1}_ocr_result.xlsx")
            wb.save(xlsx_output_path)

            print(f"Processed: {filename} (Page {page_num + 1}), results saved to {annotated_output_path} and {xlsx_output_path}")

    elif file_ext.lower() in ['.jpg', '.jpeg', '.png', '.bmp']:
        # 若是image，直接提取
        image_path = input_path
        results = ocr.ocr(image_path, cls=True, slice=slice)

        # 载入原图
        original_image = Image.open(image_path).convert("RGB")
        original_width, original_height = original_image.size

        # 创建画板
        new_image = Image.new("RGB", (original_width * 2, original_height), (255, 255, 255))

        # 同理拼接
        new_image.paste(original_image, (0, 0))

        draw = ImageDraw.Draw(new_image)

        font_path = '/System/Library/Fonts/Supplemental/Songti.ttc'
        font = ImageFont.truetype(font_path, size=70)

        for res in results:
            for line in res:
                box = [tuple(point) for point in line[0]]
                min_x = min(point[0] for point in box)
                min_y = min(point[1] for point in box)
                max_x = max(point[0] for point in box)
                max_y = max(point[1] for point in box)
                txt = line[1][0]
                score = line[1][1]

                # 绘制蓝框
                draw.rectangle([original_width + min_x, min_y, original_width + max_x, max_y], outline="blue", width=2)

                # 根据蓝框确定字体应该放在哪，目标是中间
                bbox = draw.textbbox((0, 0), txt, font=font)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                text_x = original_width + min_x + (max_x - min_x - text_width) // 2
                text_y = min_y + (max_y - min_y - text_height) // 2

                # 在框内绘制文字
                draw.text((text_x, text_y), txt, fill="black", font=font)

                # 在框底部加置信度
                draw.text((original_width + min_x, max_y + 5), f"{score:.2f}", fill="green", font=font)

                # 将数据写入Excel
                ws.append([min_x, min_y, max_x, max_y, score, txt])

        # 保存最终图片
        annotated_output_path = os.path.join(output_dir, f"{file_base}_annotated.jpg")
        new_image.save(annotated_output_path)

        # 保存为xlsx文件
        xlsx_output_path = os.path.join(output_dir, f"{file_base}_ocr_result.xlsx")
        wb.save(xlsx_output_path)

        print(f"Processed: {filename}, results saved to {annotated_output_path} and {xlsx_output_path}")

    else:
        print(f"Unsupported file: {filename}")
