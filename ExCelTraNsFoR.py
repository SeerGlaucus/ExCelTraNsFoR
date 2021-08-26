import os
import sys
import time

import engine


def print_help():
    print("**********************************文件存放位置说明**********************************")
    print("# 1.Excel文件：D:\\data\\excel                                                 ")
    print("# 2.执行脚本文件：D:\\data\\script                                               ")
    print("********************************excel命名规则说明********************************")
    print("# 1.excel文件没有强制命名规则,任意文件名称都可以                       ")
    print("# 2.excel文件与脚本中文件序号匹配,匹配规则如下                         ")
    print("# 3.excel文件列表列出来的文件名前面的[1]标号即为脚本中文件序号            ")
    print("# 4.excel文件名称变更后序号可能会变,可能会导致文件标号和脚本文件序号不匹配！ ")
    print("# 5.excel文件名变更后,脚本文件序号更改以使文件和序号匹配                 ")
    print("# 6.只支持xls文件，xlsx文件保存成xls文件可用                         ")
    print("***********************************脚本编写说明***********************************")
    print("# 1. 单元格值复制 1[1][1,1]->3[3][3,3]                             ")
    print("#    第一个1表示excel序号为1的excel,换成2的话表示序号为2的excel         ")
    print("#    第一个[1]表示第一个sheet,换成[2]的话表示第二个sheet                ")
    print("#    第一个[1,1]表示第一行第一列的单元格,换成[2,2]的话表示第二行第二列的单元格 ")
    print("#    3[3][3,3] 表示的含义与上面三行相同                               ")
    print("#    -> 表示此符号之前的为要复制的单元格,之后的为要复制到的单元格           ")
    print("#    这行脚本的含义：把第一个excel的第一个sheet的第一行第一列单元格的值复制到")
    print("#    第三个excel的第三个sheet的第三行第三类列元格上去                    ")
    print("# 2. 行复制 1[1][1,2...]->3[3][3,4...]                             ")
    print("#    2... 表示第二列及以后的列                                        ")
    print("#    这行脚本的含义：把第一个excel的第一个sheet的第一行上第二列及以后的列的值 ")
    print("#    依次复制到第三个excel的第三个sheet的第三行上从第四列开始的列上去        ")
    print("# 3. 列复制 1[1][2...,1]->3[3][3...,4]                             ")
    print("#    2... 表示第二行及以后的行                                        ")
    print("#    这行脚本的含义：把第一个excel的第一个sheet的第一列上第二行及以后的行的值 ")
    print("#    依次复制到第三个excel的第三个sheet的第四列上从第三行开始的行上去        ")
    print("# 4. 行汇总 1[1][1,2...]->3[3][3,4]                                 ")
    print("#    2... 表示第二列及以后的列                                        ")
    print("#    这行脚本的含义：把第一个excel的第一个sheet的第一行上第二列及以后的列的值 ")
    print("#    汇总填写到第三个excel的第三个sheet的第三行第四列单元格上去            ")
    print("# 5. 列汇总 1[1][2...,1]->3[3][3,4]                                ")
    print("#    2... 表示第二行及以后的行                                        ")
    print("#    这行脚本的含义：把第一个excel的第一个sheet的第一列上第二行及以后的行的值 ")
    print("#    汇总填写到第三个excel的第三个sheet的第三行第四列单元格上去             ")
    print("# 6. 单元格计算 (高级功能)                                            ")
    print("#    1[1][1,1]+2[2][2,2]->3[3][3,3]                               ")
    print("#    这行脚本的含义：两个单元格相加放到第三个单元格上去                     ")
    print("# 7. 按列搜索行单元格计算 (高级功能)                                    ")
    print("#    1[1][1[搜索内容],2]+2[2][1[搜索内容],2]->1[1][1[搜索内容],2]      ")
    print("#    这行脚本的含义：按搜索列上搜索到的内容所在行和数据列获得单元格相加放到    ")
    print("#    第三个单元格上去                                                ")
    print("# 8. 取数计算 (高级功能)                                              ")
    print("#    1[1][1...,1][1]+2[2][2...,2][2]->3[3][3...,3][3]              ")
    print("#    这行脚本比上面的每个块后面多了一部分[1],这个数字代表按照哪一列对应        ")
    print("#    这行脚本的含义：获取数据列的值,按照对应列进行+-计算,结果按照对应列放到    ")
    print("#    结果excel的对应列中                                             ")
    print("*****************************************************************************")


def main_method():
    path = "D:\\data"
    data_path = os.path.join(path, "excel")
    script_path = os.path.join(path, "script")
    data = []
    for root, dirs, files in os.walk(data_path, topdown=False):
        for name in files:
            data.append(os.path.join(root, name))
    if len(data) == 0:
        print("默认路径下（D:\\data\\excel）没有文件, 请把excel文件放到这个文件夹中,程序将在10秒内退出")
        time.sleep(10)
        sys.exit()
    print("Excel文件列表：")
    data.sort()
    a = 1
    for f in data:
        print('[' + str(a) + ']' + ' ' + f)
        a = a + 1
    script = []
    for root, dirs, files in os.walk(script_path, topdown=False):
        for name in files:
            script.append(os.path.join(root, name))
    if len(script) == 0:
        print("默认路径下（D:\\data\\script）没有文件, 请把脚本文件放到这个文件夹中,程序将在10秒内退出")
        time.sleep(10)
        sys.exit()
    print("执行脚本列表：")
    i = 1
    for sc in script:
        print('[' + str(i) + ']' + ' ' + sc)
        i = i + 1
    index = input("请选择要执行的脚本【输入数字回车即可】：")
    sc_file = script[int(index) - 1]
    print("\n开始执行脚本：" + sc_file)
    engine.call(data_path, sc_file)
    print("执行完成^_^，执行时间为：" + time.strftime('%Y/%m/%d %H:%M:%S', time.localtime(time.time())))
    print("")
    print(" ___  ___      ___      ___      ___________      ___  ___  ")
    print("|'  \/  '|    |'  \    /  '|    ('     _   ')    |'  \/'  | ")
    print(" \   \  /      \   \  //   |     )__/   \__/      \   \  /  ")
    print("  \   \/       /\   \/.    |        \_  /          \   \/   ")
    print("  /\.  \      |: \.        |        |.  |          /\.  \   ")
    print(" /  \   \     |.  \    /:  |        \:  |         /  \   \  ")
    print("|___/\___|    |___|\__/|___|         \__|        |___/\___| ")
    print("                           v0.0.1                           ")


def main():
    print_help()
    while True:
        main_method()


main()
