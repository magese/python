import os

from PIL import Image


def compress_image(input_image_path, output_image_path, quality):
    image = Image.open(input_image_path)
    if image.mode == 'RGBA':
        image = image.convert('RGB')
    image.save(output_image_path, optimize=True, quality=quality)


def main():
    image_dir = r'C:\Users\mages\Desktop\screenshot'
    image_dir2 = r'C:\Users\mages\Desktop\screenshot2'
    for root, dirs, files in os.walk(image_dir):
        for file in files:
            if file.lower().endswith(".jpg"):
                origin_path = os.path.join(root, file)
                new_path = image_dir2 + '\\' + file
                compress_image(origin_path, new_path, 80)


if __name__ == '__main__':
    main()
