# [python3-pyqt5-process-csv-data](http://www.madmalls.com/blog/post/process-csv-data-and-pyqt5/)

[![Python](https://img.shields.io/badge/Python-v3.6.1-brightgreen.svg)](https://www.python.org/)
[![PyQt5](https://img.shields.io/badge/PyQt5-v5.11.2-orange.svg)](https://www.riverbankcomputing.com/software/pyqt/intro)
[![PyInstaller](https://img.shields.io/badge/pyinstaller-v3.4-lightgrey.svg)](https://www.pyinstaller.org/)



![](http://www.madmalls.com/api/medias/uploaded/process-csv-01-54579793.png)

![](http://www.madmalls.com/api/medias/uploaded/process-csv-02-96422444.png)

![](http://www.madmalls.com/api/medias/uploaded/process-csv-03-1e3614f0.png)

![](http://www.madmalls.com/api/medias/uploaded/process-csv-04-0075221d.png)



# 1. 搭建环境

打开cmd命令行，切换到 D:\python-code\python3-pyqt5-process-csv-data 目录下

``` 
D:\python-code\python3-pyqt5-process-csv-data> python -m venv venv3
```

# 2. 激活

```
D:\python-code\python3-pyqt5-process-csv-data> venv3\Scripts\activate
(venv3) D:\python-code\python3-pyqt5-process-csv-data>
```

# 3. 安装包

```
(venv3) D:\python-code\python3-pyqt5-process-csv-data> pip install pyqt5
(venv3) D:\python-code\python3-pyqt5-process-csv-data> pip install pyinstaller
```

# 4. 图标

创建`images.qrc`，注意ico图标放在当前目录下的子目录img中：

```
<RCC>
  <qresource prefix="/" >
    <file>img/logo.ico</file>
  </qresource>
</RCC>
```

生成`images_pyqt.py`，去文件目录下执行：

```
(venv3) D:\python-code\python3-pyqt5-process-csv-data> pyrcc5 -o images_pyqt.py images.qrc
```

最后在代码中`import images_pyqt`，并且修改下图片路径，一定要在路径前面加上`冒号`:

```python
import images_pyqt

def init_ui(self):
    self.setWindowIcon(QIcon(':/img/logo.ico'))  # 图标
```

# 5. 打包成exe

```
(venv3) D:\python-code\python3-pyqt5-process-csv-data> pyinstaller --name Madman --onefile --windowed --icon=D:\python-code\python3-pyqt5-process-csv-data\logo.ico -w --paths=D:\python-code\python3-pyqt5-process-csv-data\venv3\Lib\site-packages --paths=D:\python-code\python3-pyqt5-process-csv-data pyqt5_process_csv_data.py
```