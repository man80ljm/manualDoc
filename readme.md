运行前请安装：
pip install python-docx -i https://mirrors.aliyun.com/pypi/simple/

pip install pillow==9.5.0 -i https://mirrors.aliyun.com/pypi/simple/


打包命令：
pyinstaller -F -w --add-data "C:\Windows\Fonts\simsun.ttc;." -i manualDoc.ico generate_evidence_doc.py