import datetime
import shutil
from tkinter import *
from tkinter import messagebox
from tkinter import scrolledtext
from tkinter import ttk
import tkinter.colorchooser
import os
import time
import requests
import configparser
import re
from PIL import Image
import xlrd
import windnd
import threading
import numpy as np
from cv2 import cv2


def padding(image, ksize):
    h = image.shape[0]
    w = image.shape[1]
    c = image.shape[2]

    pad = ksize // 2

    out_p = np.zeros((h + 2 * pad, w + 2 * pad, c))

    out_copy = image.copy()

    out_p[pad:pad + h, pad:pad + w, 0:c] = out_copy.astype(np.uint8)

    return out_p


def gaussian(image, ksize, sigmatmp):
    """
    1. padding
    2. 定义高斯滤波公式与卷积核
    3. 卷积过程

    高斯卷积卷积核是按照二维高斯分布规律产生的，公式为：

    G(x,y) = (1/(2*pi*(sigmatmp)^2))*e^(-(x^2+y^2)/2*sigmatmp^2)

    唯一的未知量是sigmatmp，在未指定sigmatmp的前提下，可以通过下列参考公式让程序自动选择合适的
    sigmatmp值：

    sigmatmp =  0.3 *((ksize-1)*0.5-1) + 0.8

    @ 如果mode为default，则返回abs值，否则返回unit8值
    """

    pad = ksize // 2

    out_p = padding(image, ksize)  # padding之后的图像
    # print(out_p)

    h = image.shape[0]
    w = image.shape[1]
    c = image.shape[2]

    # 高斯卷积核

    kernel = np.zeros((ksize, ksize))
    for x in range(-pad, -pad + ksize):
        for y in range(-pad, -pad + ksize):
            kernel[y + pad, x + pad] = np.exp(-(x ** 2 + y ** 2) / (2 * (sigmatmp ** 2)))
    kernel /= (sigmatmp * np.sqrt(2 * np.pi))
    kernel /= kernel.sum()

    # print(kernel)

    tmp = out_p.copy()

    # print(tmp)

    for y in range(h):
        for x in range(w):
            for z in range(c):
                out_p[pad + y, pad + x, z] = np.sum(kernel * tmp[y:y + ksize, x:x + ksize, z])

    out = out_p[pad:pad + h, pad:pad + w].astype(np.uint8)
    # print(out)

    return out


class keepCommentConfigParser():

    def __init__(self, filePath, commentPrefixes):  # commentPrefixes 需要用户输入 []，感觉这个写法不太对，最好是可以默认"#", ";" 并可以获取用户自定义字符

        self.__filePath = filePath
        self.__commentPrefixes = commentPrefixes
        self.__replacePrefix = datetime.datetime.now().strftime('%Y%m%d%H%M%S')  # 临时时间戳，解决方案的核心

        self.__oriLines = []  # 用来存储带有注释的源文件内容，保留着，后面或许有用
        self.__markupLines = []  # 用来存储用临时时间戳代替指定的注释符后的内容
        self.__backup()

        self.__updatedLines = []  # 用来存储被原 ConfigParser “践踏”后的新内容

    def __backup(self):

        # 直接复制并重命名原 .ini 文件为 .txt 文件，此处可能会有解析上的问题吧，暂未遇到
        backupFilePath = self.__replacePrefix + ".txt"
        shutil.copyfile(self.__filePath, backupFilePath)

        with open(backupFilePath, encoding="utf-8") as f:
            self.__oriLines = f.readlines()  # 获取带有注释的源文件内容
        self.__markupLines = self.__oriLines[:]  # 复制带有注释的源文件内容
        f.close()
        os.unlink(backupFilePath)  # 永久删除临时的 .txt 文件

        self.__markupComments()

    def __markupComments(self):
        if len(self.__commentPrefixes) > 0:
            for i in range(len(self.__markupLines)):
                for commentPrefix in self.__commentPrefixes:
                    if self.__markupLines[i].startswith(commentPrefix):
                        self.__markupLines[i] = self.__replacePrefix + self.__markupLines[
                            i]  # self.__markupLines 中，遇到指定注释符开头的行，则直接在行头加上临时时间戳，最终生成新的 self.__markupLines
                        break

    def update(self):

        # 读取经过原 ConfigParser “践踏”过后的 test.ini 文件
        with open(self.__filePath, encoding="utf-8") as f:
            self.__updatedLines = f.readlines()
            f.close()

        validLineId = 0
        for i in range(len(self.__markupLines)):
            # 如果 self.__markupLines 中的一行不以临时时间戳开头，则这一行直接复制“践踏”过后的 test.ini 的对应行的值（所以只能应用在修改值的场景，增减就 gg 了）
            if not self.__markupLines[i].startswith(self.__replacePrefix):
                self.__markupLines[i] = self.__updatedLines[validLineId]
                validLineId += 1
            # 如果 self.__markupLines 中的一行以临时时间戳开头，则直接去掉临时时间戳
            else:
                self.__markupLines[i] = self.__markupLines[i].replace(self.__replacePrefix, "")

                # 更新完 self.__markupLines 的所有行，重新写入文件
        with open(self.__filePath, encoding="utf-8", mode="w+") as f:
            f.writelines(self.__markupLines)
            f.close()


# 存放目录
path = os.getcwd()
print(path)
tmp_path = path + '/tmp/'
# 是否存在该目录
if not os.path.exists(tmp_path):
    os.makedirs(tmp_path)
sign_path = path + '/sign/'
# 是否存在该目录
if not os.path.exists(sign_path):
    os.makedirs(sign_path)

# 初始默认值
url = 'http://jiqie.zhenbi.com/a/re22.php'
font = 901
color = '#000000'
sigma = 0.8

# 配置文件
if not os.path.exists('config.ini'):
    # 打开一个文件,返回的是文件的文件描述符
    cfini = os.open('config.ini', os.O_CREAT | os.O_RDWR)
    # 在文件中写入内容
    print(cfini)
    initext = '[Data]\n#爬取网址\nurl = http://jiqie.zhenbi.com/a/re22.php\n#字体\nfont = 901\n#字体颜色\n' \
              'color = #000000\n#平滑系数,范围大于等于0,小数点一位,例如0.8,1.5\n#平滑系数设置为 0 时,不进行平滑处理\n' \
              '#平滑系数越大,看起来越平滑,但相应会看起来模糊些\nsigma = 0.8\n'
    os.write(cfini, initext.encode('utf-8'))
    # 关闭文件
    os.close(cfini)

kCCP = keepCommentConfigParser('config.ini', ["#", ";"])
cf = configparser.ConfigParser(allow_no_value=True)  # 这个 allow_no_value = True 貌似是重点
cf.read('config.ini', encoding="utf-8")
# cf = configparser.ConfigParser()
# 读取配置文件，如果写文件的绝对路径，就可以不用os模块
# cf.read(path + '/config.ini', encoding='utf-8')
# 获取文件中所有的section(一个配置文件中可以有多个配置，如数据库相关的配置，邮箱相关的配置，
# 每个section由[]包裹，即[section])，并以列表的形式返回
secs = cf.sections()
# 获取某个section名为Data所对应的键
# options = cf.options('Data')
# print(options)
# 获取section名为Mysql-Database所对应的全部键值对
# items = cf.items('Data')
# print(items)
# 获取[Mysql-Database]中host对应的值
url = cf.get('Data', 'url')
font = int(cf.get('Data', 'font'))
color = cf.get('Data', 'color')
sigma = float(cf.get('Data', 'sigma'))


# 检查路径是否存在
def checkpathstate():
    # 是否存在该目录
    if not os.path.exists(tmp_path):
        os.makedirs(tmp_path)
    # 是否存在该目录
    if not os.path.exists(sign_path):
        os.makedirs(sign_path)


def thread_it(func, *args):
    # '''将函数打包进线程'''
    # 创建
    t = threading.Thread(target=func, args=args)
    # 守护 !!!
    t.setDaemon(True)
    # 启动
    t.start()
    # 阻塞--卡死界面！
    # t.join()


# 定时
def sleep(seconds):
    start_time = time.time()  # 此处返回python调用函数开始时的时间戳
    while time.time() - start_time < seconds:  # 一直在比较是否小于seconds，直到大于或等于seconds停止
        pass


# 拖拽excel到程序中
def dragged_files(files):
    try:
        msg = '/n'.join((item.decode('gbk') for item in files))
        print(msg)
        # read_excel(msg)
        thread_it(read_excel, msg)
    except BaseException:
        pass
        # tkinter.messagebox.showwarning('警告', '请正确操作软件!')


# 读取excel
def read_excel(excelPath):
    try:
        # 打开文件
        workBook = xlrd.open_workbook(excelPath)
        # print(workBook)
    except BaseException:
        tkinter.messagebox.showwarning('警告', '请拖入正确的excel!')
    else:
        res = '准备处理>>  \n'
        dhTextK.config(state='normal')
        dhTextK.insert('end', res)
        dhTextK.config(state='disabled')
        # 1.获取sheet的名字
        # 1.1 获取所有sheet的名字(list类型)
        allSheetNames = workBook.sheet_names()
        # 1.2 按索引号获取sheet的名字（string类型）
        sheet1Name = workBook.sheet_names()[0]
        print(sheet1Name)

        # 2. 获取sheet内容
        # # 2.1 法1：按索引号获取sheet内容
        sheet1_content1 = workBook.sheet_by_index(0)  # sheet索引从0开始
        # # 2.2 法2：按sheet名字获取sheet内容
        sheet1_content2 = workBook.sheet_by_name('Sheet1')

        # 3. sheet的名称，行数，列数
        print(sheet1_content1.name, sheet1_content1.nrows, sheet1_content1.ncols)

        # 循环获取编号和数据
        rows = sheet1_content1.nrows
        for i in range(rows):
            index = str(int(sheet1_content1.row_values(i)[0]))
            name = sheet1_content1.row_values(i)[1]
            if name:
                # thread_it(sleep, 1)
                # thread_it(print, i)
                sleep(1)
                # print(i)
                # 检查目录
                checkpathstate()
                signAuto(index, name, font, color)
                print(index)
                print(name)
            # else:
            #     break
        dhTextK.config(state='normal')
        dhTextK.insert('end', '完成 \n即将退出 \n')
        dhTextK.config(state='disabled')
        dhTextK.see(tkinter.END)
        sleep(1)
        root.quit()


# 将gif图片转成PNG图片
def iter_frames(im):
    try:
        i = 0
        while 1:
            im.seek(i)
            imframe = im.copy()
            if i == 0:
                palette = imframe.getpalette()
            else:
                imframe.putpalette(palette)
            yield imframe
            i += 1
    except BaseException:
        pass
        # tkinter.messagebox.showwarning('警告', '请正确操作软件!111')


# 高斯模糊 平滑锯齿
def pinghuapng(pathtmp):
    global sigma
    if sigma == 0:
        print(sigma)
        return
    else:
        img = cv2.imread(pathtmp, cv2.IMREAD_UNCHANGED)
        gaussian_img = gaussian(img, 3, float(sigma))
        print(gaussian_img)
        cv2.imwrite(pathtmp, gaussian_img, [cv2.IMWRITE_PNG_COMPRESSION, 9])


# 定义爬虫函数
def signAuto(index, name, font, color):
    # 参数
    data = {
        'id': name,
        'zhenbi': 20191123,
        'id1': 800,
        'id2': font,
        'id3': color,
        'id4': color,
        'id5': color,
        'id6': color
    }
    try:
        # 抓取图片
        result = requests.post(url, data=data)
        result.encoding = 'utf-8'
        html = result.text
        reg = '<img src="(.*?)">'
        img_path = re.findall(reg, html)
        # 图片完整路径
        img_url = img_path[0]
        print(img_url)
    except:
        tkinter.messagebox.showwarning('警告', '爬虫程序出错了,请检查网络后重试!')
    else:
        try:
            # 获取图片内容
            response = requests.get(img_url).content
            f = open(tmp_path + 'sign_' + '{}.gif'.format(index), 'wb')
            # 写入
            f.write(response)
            f.close()
        except:
            tkinter.messagebox.showwarning('警告', '图片写入失败,请检查sign文件夹和tmp文件夹是否存在!')
        else:
            try:
                # 将gif图片转成PNG图片
                im = Image.open(tmp_path + 'sign_' + str(index) + '.gif')
                for i, frame in enumerate(iter_frames(im)):
                    frame.save(sign_path + 'sign_' + str(index) + '.png', **frame.info)
            except:
                tkinter.messagebox.showwarning('警告', '图片处理失败,请重试!')
            else:
                try:
                    # 高斯模糊 平滑锯齿
                    pathtmp = sign_path + 'sign_' + str(index) + '.png'
                    pinghuapng(pathtmp)
                except:
                    tkinter.messagebox.showwarning('警告', '平滑处理失败,请重试!')
                else:
                    # 显示进度
                    res = '处理完成>>  编号: ' + str(index) + ' 姓名: ' + name + '\n'
                    dhTextK.config(state='normal')
                    dhTextK.insert('end', res)
                    dhTextK.config(state='disabled')
                    dhTextK.see(tkinter.END)


# 颜色选择
def chooseColor():
    try:
        global color, colorTextVar
        r = tkinter.colorchooser.askcolor(title='颜色选择器')
        tmp = str(r[1])
        if tmp == 'None':
            color = color
        else:
            color = r[1]
        btn2.config(bg=color)
        colorTextVar.set(color)
        print(r[1])
        try:
            # 保存配置信息
            kCCP = keepCommentConfigParser('config.ini', ["#", ";"])
            cf = configparser.ConfigParser(allow_no_value=True)  # 这个 allow_no_value = True 貌似是重点
            cf.read('config.ini', encoding="utf-8")
            cf.set('Data', 'color', color)
            with open('config.ini', 'w+') as f:
                cf.write(f)
            f.close()
            kCCP.update()
        except:
            tkinter.messagebox.showwarning('警告', '保存颜色配置信息失败,对程序无影响!')
    except BaseException:
        tkinter.messagebox.showwarning('警告', '颜色选择器莫名出错,请重开!')
        root.quit()


# 下拉选择
def comboboxChoose(event):
    try:
        global font
        if combobox.get() == '一笔艺术签':
            font = 901
        elif combobox.get() == '连笔商务签':
            font = 904
        else:
            font = 905
        print(combobox.get())
        try:
            # 保存配置信息
            kCCP = keepCommentConfigParser('config.ini', ["#", ";"])
            cf = configparser.ConfigParser(allow_no_value=True)  # 这个 allow_no_value = True 貌似是重点
            cf.read('config.ini', encoding="utf-8")
            cf.set('Data', 'font', str(font))
            with open('config.ini', 'w+') as f:
                cf.write(f)
            f.close()
            kCCP.update()
        except:
            tkinter.messagebox.showwarning('警告', '保存字体配置信息失败,对程序无影响!')
    except BaseException:
        tkinter.messagebox.showwarning('警告', '下拉选择框意外出错,请重选!')


# 检测输入颜色格式
def colorTextVarTest(event):
    try:
        global color
        root.focus_set()
        text = colorTextVar.get()
        arrs = list(text)
        print(arrs)
        if arrs[0] == '#' and len(arrs) == 7:
            del arrs[0]
            str = ''.join(arrs)
            print(str)
            res = re.match('[0-9a-fA-F]{6}', str)
            print(res)
            if res:
                color = text
                btn2.config(bg=color)
                try:
                    # 保存配置信息
                    kCCP = keepCommentConfigParser('config.ini', ["#", ";"])
                    cf = configparser.ConfigParser(allow_no_value=True)  # 这个 allow_no_value = True 貌似是重点
                    cf.read('config.ini', encoding="utf-8")
                    cf.set('Data', 'color', color)
                    with open('config.ini', 'w+') as f:
                        cf.write(f)
                    f.close()
                    kCCP.update()
                except:
                    tkinter.messagebox.showwarning('警告', '保存字体配置信息失败,对程序无影响!')
            else:
                tkinter.messagebox.showwarning('警告', '颜色值错误!')
                colorTextVar.set(color)
        else:
            tkinter.messagebox.showwarning('警告', '颜色值错误!')
            colorTextVar.set(color)
    except BaseException:
        tkinter.messagebox.showwarning('警告', '请正确输入颜色值!')
        colorTextVar.set(color)


# 检测输入平滑系数
def pingHuaXiShuVarTest(event):
    global sigma
    root.focus_set()
    num = pingHuaXiShuVar.get()
    try:
        n1 = eval(num)
        n2 = float(num)
    except:
        tkinter.messagebox.showwarning('警告', '平滑系数错误!\n范围大于等于0,例如 0.5')
        pingHuaXiShuVar.set(sigma)
    else:
        sigma = n2
        try:
            # 保存配置信息
            kCCP = keepCommentConfigParser('config.ini', ["#", ";"])
            cf = configparser.ConfigParser(allow_no_value=True)  # 这个 allow_no_value = True 貌似是重点
            cf.read('config.ini', encoding="utf-8")
            cf.set('Data', 'sigma', str(sigma))
            with open('config.ini', 'w+') as f:
                cf.write(f)
            f.close()
            kCCP.update()
        except:
            tkinter.messagebox.showwarning('警告', '保存字体配置信息失败,对程序无影响!')


# 创建窗口
root = Tk()
# 标题
root.title('签名自动化      by小木')
# 窗口大小
root.geometry('700x320')
# 窗口的初始位置
root.geometry('+400+200')
# #### 颜色选择器 #####
# 按钮
btn1 = Button(root, text='字体颜色>>', font=('宋体', 16), state=DISABLED, bd=0)
btn2 = Button(root, text='         ', font=('宋体', 16), bg=color, bd=0, command=chooseColor, width=10)
btn4 = Button(root, text='字体颜色>>', font=('宋体', 16), state=DISABLED, bd=0)
btn44 = Button(root, text='<<输完回车', font=('宋体', 16), state=DISABLED, bd=0)
btn5 = Button(root, text='平滑系数>>', font=('宋体', 16), state=DISABLED, bd=0)
btn55 = Button(root, text='<<输完回车', font=('宋体', 16), state=DISABLED, bd=0)
# 设置按钮的位置
btn1.grid(row=0, column=0)
btn2.grid(row=0, column=1, sticky=W)
btn4.grid(row=1, column=0)
btn44.grid(row=1, column=1)
btn5.grid(row=2, column=0)
btn55.grid(row=2, column=1)
# 输入框
colorTextVar = StringVar()
colorTextVar.set(color)
colorText = Entry(root, font=('宋体', 16), textvariable=colorTextVar, width=10)
colorText.bind('<Return>', colorTextVarTest)
# 设置输入框的位置
colorText.grid(row=1, column=1, sticky=W)
# 输入框
pingHuaXiShuVar = StringVar()
pingHuaXiShuVar.set(sigma)
pingHuaXiShu = Entry(root, font=('宋体', 16), textvariable=pingHuaXiShuVar, width=10)
pingHuaXiShu.bind('<Return>', pingHuaXiShuVarTest)
# 设置输入框的位置
pingHuaXiShu.grid(row=2, column=1, sticky=W)
# #### 颜色选择器 #####
# #### 字体选择 #####
# 按钮
button3 = Button(root, text='字体选择>>', font=('宋体', 16), state=DISABLED, bd=0)
# 下拉框
combobox = ttk.Combobox(root, font=('宋体', 16), width=15)
# 设置下拉菜单中的值
combobox['value'] = ("一笔艺术签", "连笔商务签", "一笔商务签")
combobox['state'] = "readonly"  # 设定下拉框状态，readonly表示只读，不可更改内容
# 设置下拉菜单的默认值,默认值索引从0开始
combobox.current(0)
combobox.bind("<<ComboboxSelected>>", comboboxChoose)  # <ComboboxSelected>当列表选择时触发绑定函数
# 设置下拉框位置
button3.grid(row=3, column=0)
combobox.grid(row=3, column=1, sticky=W)
# #### 字体选择 #####
# 状态
# 按钮
button4 = Button(root, text='    状态>>', font=('宋体', 16), state=DISABLED, bd=0)
# 设置按钮的位置
button4.grid(row=4, column=0)
# 多行文本框
# 滚动文本框
scrolW = 40
scrolH = 5
# wrap=tk.WORD   这个值表示在行的末尾如果有一个单词跨行，会将该单词放到下一行显示,
# 比如输入hello，he在第一行的行尾,llo在第二行的行首,
# 这时如果wrap=tk.WORD，则表示会将 hello 这个单词挪到下一行行首显示, wrap默认的值为tk.CHAR
dhTextK = scrolledtext.ScrolledText(root, font=('宋体', 16), width=scrolW, height=scrolH, wrap=tkinter.WORD,
                                    state='disabled')
dhTextK.grid(row=4, column=1)
dhTextK.config(state='normal')
dhTextK.insert('end', '等待拖入excel后自动执行 \n')
dhTextK.config(state='disabled')
# 拖拽
windnd.hook_dropfiles(root, func=dragged_files)
# 显示窗口
root.mainloop()
