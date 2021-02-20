import re
import os


def remove_duplicate(file, file2):
    list01 = []
    text = open(file, 'r', encoding='UTF-8')
    for line in text:   # 需要在最后一行添加换行符
        if line in list01:
            continue
        list01.append(line)
    with open(file2, "w", encoding='UTF-8') as s:
        s.writelines(list01)
    return


#读取目录下的所有文件，包括嵌套的文件夹
def GetFileList(dir, fileList):
    newDir = dir
    if os.path.isfile(dir):
        fileList.append(dir)
    elif os.path.isdir(dir):
        for s in os.listdir(dir):
            # 如果需要忽略某些文件夹，使用以下代码
            # if s == "xxx":
            # continue
            newDir = os.path.join(dir, s)
            GetFileList(newDir, fileList)
    return fileList


def extract_text(file):
    list1 = GetFileList(file, [])
    # 打开一个文件
    fo = open("E:\\TestDatas\\extract.txt", "w")  # 打开文件
    for i in list1:
        print(i)  # 测试完整文件路径
        # print(os.path.basename(i))  # 文件名
        # index = i.find(".", 0)  # 找到点号的位置
        # print(i[index - 1:index])  # 截取目标字符
        # print(os.path.basename(i) + " " + i[index - 1:index])  # 测试目标字符串
        # fo.write(os.path.basename(i) + " " + i[index - 1:index] + "\n")  # 将目标字符串写入文件
        remove_duplicate(i, fo)
    fo.close()  # 关闭打开的文件


def remove_text(file, file2):
    list01 = []
    text = open(file, 'r', encoding='UTF-8')
    for line in text:   # 需要在最后一行添加换行符
        # 提取，，后面的字符
        index = line.find(",,")
        line2 = line.replace(line[0:index+2], "")

        list01.append(line2)
    with open(file2, "w", encoding='UTF-8') as s:
        s.writelines(list01)
    return


if __name__ == '__main__':
    # extract_text('D:\\MES\\IM\\取り出し\\OutPut\\')
    remove_text('E:\\TestDatas\\extract.txt', 'E:\\TestDatas\\text2.txt')
    remove_duplicate('E:\\TestDatas\\text2.txt', 'E:\\TestDatas\\POEM_out.txt')
