import base64
import os


def trans_base64_1(file_path):
    with open(file_path, "rb") as f:
        return str(base64.b64encode(f.read()))[2:-1]


filepath = r'C:\Users\mages\Desktop\favicon.ico'
print(os.path.abspath(os.path.dirname(__file__)))
