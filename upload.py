# Python3脚本，不适用于Python2
# !/usr/bin/envpython
# coding=utf-8
from bottle import route, run, request, static_file
import os
import getpass
import sys

operating_system = sys.platform
user = getpass.getuser()
print("Current system is", operating_system)
print("Current user is ", user)
xlsx_path = r"/Users/" + user + "/healthcodeGHC"
if operating_system == "darwin":
    xlsx_path = r"/Users/" + user + "/healthcodeGHC"
elif operating_system ==  "linux":
    xlsx_path = r"/home/" + user + "/healthcodeGHC"
elif operating_system == "win32":
    xlsx_path = "C:\\Users\\" + user + "\\healthcodeGHC"
elif operating_system == "cygwin":
    xlsx_path = "C:\\Users\\" + user + "\\healthcodeGHC"
else:
    print("Unknown system.")

# 此处可扩充为完整HTML
uploadPage = '''
    <body id="tinymce" class="mce-content-body " data-id="content" contenteditable="true" spellcheck="false">
        <h1> 注意事项</h1>
        <h3> 
            <ol>
                <li>上传文件必须是粤省事导出来的Excel文件</li>
                <li>修改文件名，不要有特殊符号，最好改成日期，如503-20220606.xlsx</li>
                <li>点击上传后，请耐心等待返回<span style="color: rgb(224, 62, 45);" data-mce-style="color: #e03e2d;">下载文件</span>。
                </li>
                <li>保持本页面不要关闭。</li>
                <li>直到本页面出现<span style="color: rgb(224, 62, 45);" data-mce-style="color: #e03e2d;">下载文件</span>后即可点击下载。
                </li>
                <li>若要再次执行，请重新打开本网址。</li>
            </ol>
            <form action="upload" method="POST" enctype="multipart/form-data">
                <input type="file" name="data" />
                <input type="submit" value="上传" />
            </form>
	    </h3>
    </body>
'''


@route('/upload')
def upload():
    return uploadPage


@route('/upload', method='POST')
def do_upload():
    upload_file = request.files.get('data')  # 获取上传的文件
    upload_file.save(xlsx_path, overwrite=True)  # overwrite参数是指覆盖同名文件
    if file_filter(upload_file.filename):
        if os.system('python main.py %s' % upload_file.filename) == 0:
            output_file = '学生_' + upload_file.filename
            return u"<h1>过滤成功，请点击<a href='/download/" + output_file + "'>下载文件</a>, 或者<a href='/upload'>返回首页</a></h1>"
        else:
            return u"<h1>出错了！请检查上传的文件或者联系管理员！</h1>"
    else:
        return u"<h1>出错了！请检查上传的文件或者联系管理员！</h1>"


@route('/download/<filename:path>')
def download(filename):
    return static_file(filename, root=xlsx_path, download=filename)


def file_filter(f):
    if f[-5:] in ['.xlsx']:
        return True
    return False


run(host='0.0.0.0', port=8899, debug=False, server='cheroot')
