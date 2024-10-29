import os

from PIL import Image


def compress_image(input_image_path, output_image_path, quality):
    image = Image.open(input_image_path)
    if image.mode == 'RGBA':
        image = image.convert('RGB')
    image.save(output_image_path, optimize=True, quality=quality)

def cut_image(input_image_path, output_image_path, box):
    image = Image.open(input_image_path)
    crop_image = image.crop(box)
    crop_image.save(output_image_path)


def main():
    image_dir = r'C:\Users\Magese\Desktop\screenshot'
    image_dir2 = r'C:\Users\Magese\Desktop\screenshot2'
    for root, dirs, files in os.walk(image_dir):
        for file in files:
            if file.lower().endswith(".jpg"):
                # 文件路径
                origin_path = os.path.join(root, file)
                new_path = image_dir2 + '\\' + file
                # 裁剪图片
                image = Image.open(origin_path)
                width, height = image.size
                box = (370, 80, width - 100, height)
                crop_image = image.crop(box)
                crop_image.save(new_path)
                print('cut image success, output path:', new_path)
                # 压缩图片
                compress_image(new_path, new_path, 90)


if __name__ == '__main__':
    main()
