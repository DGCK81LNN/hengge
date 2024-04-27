# -*- coding: utf-8 -*-
import os
import wget
import yaml
from pynput import keyboard
from time import sleep
from traceback import print_exc
from urllib.parse import quote as urlquote
path = os.path

COPY_TYPES = {
    "word": "DOC",
    "excel": "XLS",
    "ppt": "PPT",
}
COPY_DEST_NAMES = {
    "word": "文档.docx",
    "excel": "文档.xlsx",
    "ppt": "文档.PPTX",
}

def dazi(kc, text, percentage=100):
    print("请回到考试系统点击“开始打字”，程序即将开始自动输入！")
    for i in range(5, 0, -1):
        print(f"<<[    {i}    ]>>", end="\r")
        sleep(1)
    print("<<[    0    ]>>")
    n = round(len(text) * (percentage / 100.0))
    for i in range(n):
        kc.type(text[i])
        sleep(0.012)

def choose_task(tasks):
    print()
    print("(0) 返回")
    for i, task in enumerate(tasks):
        print(f"({i + 1}) {task['name']}")
    print()
    try:
        task_index = int(input("输入任务序号> ")) - 1
        print()
        if task_index == -1:
            return True
        elif task_index < -1 or task_index >= len(tasks):
            raise ValueError()
    except ValueError:
        print("输入有误！")
        return False
    return tasks[task_index]

def func_dazi():
    kc = keyboard.Controller()
    with open("dazi.yml", "r", encoding="utf-8") as f:
        tasks = yaml.load(f, Loader=yaml.FullLoader)
    while True:
        # 选择任务
        task = choose_task(tasks)
        if task is True: break
        elif task is False: continue
        texts = task['texts']

        while True:
            # 选择题目
            print()
            print(" (0) 返回")
            for i, text in enumerate(texts):
                print()
                print("%4s" % f"({i + 1})" + f" {text[44:80]}…")
            print()
            try:
                text_index = int(input("输入打字题号> ")) - 1
                print()
                if text_index == -1:
                    break
                elif text_index < -1 or text_index >= len(texts):
                    raise ValueError()
            except ValueError:
                print("输入有误！")
                continue
            print()
            text = texts[text_index]

            # 输入目标正确率
            try:
                percentage = float(input("想要程序给你打到多少正确率？输入（0, 100］的百分比。输入 0 取消> "))
                print()
                if percentage == 0:
                    continue
                elif percentage < 0 or percentage > 100:
                    raise ValueError()
            except ValueError:
                print("输入有误！")
                continue

            dazi(kc, text, percentage)

            if len(texts) == 1:
                break

def func_copy():
    with open("copy.yml", "r", encoding="utf-8") as f:
        tasks = yaml.load(f, Loader=yaml.FullLoader)

    # 定位试题文件夹
    print()
    print("注意，请先在考试系统中开始任务，再启动本程序！")
    dirs = [file for file in os.scandir("D:\\Exam") if file.is_dir()]
    dir = max(*dirs, key=os.path.getmtime).path
    print(f"当前试题文件夹：{dir}")

    while True:
        # 选择任务
        task = choose_task(tasks)
        if task is True: break
        elif task is False: continue

        # 确定下载地址和保存路径
        downloads = []
        for type in COPY_TYPES:
            if type not in task: continue
            src_names = task[type]
            type_path = path.join(dir, COPY_TYPES[type])
            if not path.isdir(type_path):
                print("错误：文档类型不一致，请检查所选任务是否正确，是否已在考试系统中开始任务！")
                downloads.clear()
                break
            dest_paths = [path.join(ent.path, COPY_DEST_NAMES[type]) for ent in os.scandir(type_path) if ent.is_dir()]
            if len(src_names) != len(dest_paths):
                print("错误：文档数量不一致，请检查所选任务是否正确，是否已在考试系统中开始任务！")
                downloads.clear()
                break
            dest_paths.sort(key=lambda x: (len(x), x))
            for i, src_name in enumerate(src_names):
                downloads.append((src_name, dest_paths[i]))
        if len(downloads) == 0:
            input("按回车键返回> ")
            continue

        # 复制文件
        for src_name, dest_path in downloads:
            url = f"https://vudrux.site/d/docs/{src_name}"
            print(url, "-->", dest_path)
            if path.exists(dest_path):
                os.unlink(dest_path)
            wget.download(url, dest_path)
            print()
    
        print("复制成功！")

def try_download_data_file(url, name):
    try:
        dest = wget.download(url, name)
        if dest != name:
            os.unlink(name)
            os.rename(dest, name)
    except Exception:
        print_exc()
    print()

funcs = [
    ("自动打字", func_dazi),
    ("复制哼歌文件", func_copy),
]

print("哼歌 v1.0.1")
print("更新数据…")
try_download_data_file("https://vudrux.site/d/docs/copy.yml", "copy.yml")
try_download_data_file("https://vudrux.site/d/docs/dazi.yml", "dazi.yml")

while True:
    # 选择功能
    print()
    print("(0) 退出")
    for i, func in enumerate(funcs):
        print(f"({i + 1}) {func[0]}")
    print()
    try:
        func_index = int(input("输入要使用的功能序号> ")) - 1
        print()
        if func_index == -1:
            break
        elif func_index < 0 or func_index >= len(funcs):
            raise ValueError()
    except ValueError:
        print("输入有误！")
        continue
    try:
        funcs[func_index][1]()
    except (EOFError, KeyboardInterrupt):
        break
    except Exception:
        print_exc()
        print("抱歉，程序运行出错！")
