# -*- encoding:utf-8 -*-

import pandas as pd
import datetime
import os


path = os.path.abspath('.')

df = pd.read_excel(os.path.join(path, '采样.xlsx'),
                   engine='openpyxl', header=0)  # 获取当前路径
time = str(datetime.datetime.now().strftime('%H-%M'))
date = str(datetime.datetime.now().strftime('%m/%d'))
filename = str('浦东医院'+time+'汇报'+'.txt')
title_name = df.columns.values.tolist()
point = df["送检单位"].values.tolist()   # 数组
point_original = list(point)  # 数组转换为列表
total_index = title_name.index("总数量")
operat_index = title_name.index("已扩增")  # 获取上机扩增下标值
finish_index = title_name.index("检测完成")  # 获取检测完成的下标值
source_index = title_name.index("送检批次")
# 提取需总送样量及上机数量，检测完成数


def total():
    operat_num = 0
    total_num = 0
    finish_num = 0
    wait_num = 0

    for i, sample_point in enumerate(point):
        if '紧急送样' in sample_point:
            total_num = int(df.iloc[i, total_index]) + total_num
            operat_num = int(df.iloc[i, operat_index]) + operat_num
            finish_num = int(df.iloc[i, finish_index]) + finish_num
    wait_num = total_num - finish_num
    output.write(("汇报时间{4}\n医院名称：浦东医院\n共收到紧急样本数{0}份\n清点完毕已上机扩增{1}份\n核酸检测已出报告{2}份\n阴性报告{2}份\n待出报告{3}份\n--------------\n\n".format(
        total_num, operat_num, finish_num, wait_num, time)))
    # output.close()

# 来源明确，紧急送样的标本


def source():
    # output = open(os.path.join(path,'output.txt'),'a')
    # point = df["送检单位"].values.tolist()
    for i, sample_point in enumerate(point):
        if "紧急送样" in sample_point:
            total_num = int(df.iloc[i, total_index])
            operat_num = int(df.iloc[i, operat_index])
            finish_num = int(df.iloc[i, finish_index])
            wait_num = total_num - finish_num
            source = df.iloc[i, source_index]
            output.write("其中{4}样本数{0}份\n清点完毕已上机扩增{1}份\n核酸检测已出报告{2}份\n阴性报告{2}份\n待出报告{3}份\n--------------\n".format(
                total_num, operat_num, finish_num, wait_num, source))
    output.write("\n\n=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=\n\n")

    # output.close()


'''
#未知来源（紧急送样）
def unknow():
    #output = open(os.path.join(path,'output.txt'),'a')
    point = df["送检单位"].values.tolist()
    unknowtotal_num = 0
    unknowoperat_num=0
    unknowfinish_num=0
    for i ,sample_point in enumerate(point):
        if "紧急送样" in sample_point:
            if "未知"  in sample_point:
                unknowtotal_num = int(df.iloc[i,total_index])+ unknowtotal_num
                unknowoperat_num = int(df.iloc[i,operat_index]) + unknowoperat_num
                unknowfinish_num = int(df.iloc[i,finish_index]) + unknowfinish_num
    unknowwait_num = unknowtotal_num - unknowfinish_num
    #print("其中未知来源样本数{0}份\n清点完毕已上机扩增{1}份\n核酸检测已出报告{2}份\n阴性报告{2}份\n待出报告{3}份\n".format(unknowtotal_num,unknowoperat_num,unknowfinish_num,unknowwait_num))
    output.write("其中未知来源样本数{0}份\n清点完毕已上机扩增{1}份\n核酸检测已出报告{2}份\n阴性报告{2}份\n待出报告{3}份\n\n=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=\n\n\n".format(unknowtotal_num,unknowoperat_num,unknowfinish_num,unknowwait_num))
    #output.close()
'''
# 社区统计


def community():
    # output = open(os.path.join(path,'output.txt'),'a')
    # point = df["送检单位"].values.tolist()
    new_list_commulity = set()
    comm_total_num = 0  # 在社区列表内的检测标本数量统计
    allcomm_total_num = 0  # 所有检测标本统计

    list_commulity = {'六灶社区', '宣桥社区', '万祥社区', '泥城社区', '书院社区', '南华医院', '大团社区', '惠南社区', '老港社区', '芦潮港社区', '隔离点',
                      '紧急送样', '大棚云', '大棚', '急诊', '发热门诊', '隔离病房', '南精卫', '观察区', '筛查', '祝桥社区', '住院', '25号点', '44号点', '9号点'}

    # 创建当天有数据的社区列表
    for commulity in list_commulity:
        for sample_point in point:
            if commulity == sample_point:
                new_list_commulity.add(commulity)  # 获取到当天有数据的社区列表

    for commulity in new_list_commulity:
        for i, sample_point in enumerate(point):
            if commulity == sample_point:
                point_num = int(df.iloc[i, total_index])
                comm_total_num = point_num + comm_total_num  # 相同社区不同批次的数据相加
        output.write("{0}:{1}人份\n".format(commulity, comm_total_num))
        allcomm_total_num = allcomm_total_num + comm_total_num  # 不同社区数据汇总
        comm_total_num = 0

    # 删除在社区列表list_commulity中的检测标本，汇总除此之外的数据
    del_list = {'总合计'}  # 创建需要删除送检单位的不重复数据合集
    for commulity in list_commulity:
        for sample_point in point:
            if commulity in sample_point:
                del_list.add(sample_point)

    for position in del_list:
        for i in range(len(point)-1, -1, -1):  # 倒序删除，可防止出现重复值及下标的自动转换
            if point[i] == position:
                point.pop(i)
    for sample_point in point:
        for i in range(len(point_original)-1, -1, -1):
            if point_original[i] == sample_point:
                point_num = int(df.iloc[i, total_index])
                comm_total_num = point_num+comm_total_num
        allcomm_total_num = allcomm_total_num + comm_total_num
        output.write("{0}:{1}人份\n".format(sample_point, comm_total_num))
        comm_total_num = 0
    output.write("{0}检验科核酸自测共计{1}份\n\n\n".format(date, allcomm_total_num))
    # output.close()


if __name__ == '__main__':
    while True:
        output = open(os.path.join(path, filename), 'w')
        total()
        source()
        # unknow()
        community()
        output.close()
        break
