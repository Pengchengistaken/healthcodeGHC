# 用途
* 粤省事APP推出团体码后，学校可以以班级为单位直接收集学生及其同住人核酸检测信息。
* 见公众号文章：https://mp.weixin.qq.com/s/ApZbtf_YiCu8xrpMdH3KCg
* 有的学校要求收集学生及同住人的健康信息
* 有的学校要求只收集学生的健康信息
* 本程序用于排除掉粤省事中团体码中已经存在的同住人的信息，只显示团体码内本人的信息。

# 输入
粤省事APP导出来的所有人的健康信息Excel表格。

# 输出
将本团体内的人（本人）的健康信息筛选出来，剔除掉添加的家属。

# 环境安装
```commandline
pip install -r requirements.txt
```

# 使用方法
* 启用用于上传Excel表格的网站，默认本地IP。或者自行搭建端口转发/URL域名。
```commandline
python upload.py
```
* 网址为：
```commandline
http://your_ip:8899/upload
```
* 粤省事APP上导出Excel表格。
* 打开本网站，点击上传将表格上传。
* 筛选后点击下载，下载的表格即过滤了同住人的信息，并将健康信息标志不同的颜色。