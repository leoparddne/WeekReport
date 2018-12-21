# -*- coding: utf-8 -*-
"""
12.22 2:12AM
"""
#import time
import datetime
import xlwt

now=datetime.datetime.now()
workbook=xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('Sheet1')
count=1#序号
newDay=1#是否为另一天的记录
f=open(r"C:\Users\Dell\Desktop\a.txt")
i=4#行号
hour=6#工时
context=0
ids=1
thing=2#工作事项
project=3#所属项目
skill=4#技术点分解描述
other=5#需要协调内容（多人接口联调、现场测试、客户试用等等）
planStart=7#计划开始时间
planEnd=8#计划结束时间
start=9#时间开始时间
end=10#时间完成时间
statu=11#完成状态
backup=12#总结备案

#样式
borders = xlwt.Borders()  # Create borders
 
borders.left = xlwt.Borders.THIN  # 添加边框-虚线边框
borders.right = xlwt.Borders.THIN  # 添加边框-虚线边框
borders.top = xlwt.Borders.THIN  # 添加边框-虚线边框
borders.bottom = xlwt.Borders.THIN  # 添加边框-虚线边框
borders.left_colour = 0x08 # 边框上色
borders.right_colour = 0x08
borders.top_colour = 0x08
borders.bottom_colour = 0x08

#列宽
worksheet.col(context).width=256*10
worksheet.col(ids).width=256*8
worksheet.col(thing).width=256*30
worksheet.col(project).width=256*10
worksheet.col(skill).width=256*20
worksheet.col(other).width=256*25
worksheet.col(hour).width=256*10
worksheet.col(planStart).width=256*25
worksheet.col(planEnd).width=256*25
worksheet.col(start).width=256*25
worksheet.col(end).width=256*25
worksheet.col(statu).width=256*12
worksheet.col(backup).width=256*22
#创建一个样式----------------------------
stylefirstRow = xlwt.XFStyle() 
stylefirstRow.borders = borders
pattern = xlwt.Pattern() 
pattern.pattern =xlwt. Pattern.SOLID_PATTERN 
pattern.pattern_fore_colour = xlwt.Style.colour_map['gold'] 
#设置单元格背景色为黄色
stylefirstRow.pattern = pattern
al = xlwt.Alignment()
al.horz = 0x02      # 设置水平居中
al.vert = 0x01      # 设置垂直居中
al.wrap=1    #自动换行
stylefirstRow.alignment = al
font = xlwt.Font() # 为样式创建字体
#font.name = 'Times New Roman' 
font.name = u'微软雅黑' 
#font.bold = True # 黑体
#font.underline = True # 下划线
#font.italic = True # 斜体字
font.height=0x00C8*1.6 # C8 in Hex (in decimal) = 10 points in height.
#font.colour_index=48#蓝色
#font.colour_index=2#红色
#font.colour_index=59#灰色
#font.colour_index=44#浅蓝色
stylefirstRow.font = font # 设定样式
#第二行样式----------------------------
stylesecondRow = xlwt.XFStyle() 
stylesecondRow.borders = borders
font2 = xlwt.Font() # 为样式创建字体
font2.name = u'微软雅黑' 
font2.height=0x00C8*1.4 # C8 in Hex (in decimal) = 10 points in height.
stylesecondRow.font = font2 # 设定样式
stylesecondRow.alignment = al
#第3行样式----------------------------
style3Row = xlwt.XFStyle() 
style3Row.borders = borders
font3 = xlwt.Font() # 为样式创建字体
font3.name = u'微软雅黑' 
font3.height=0x00C8 # C8 in Hex (in decimal) = 10 points in height.
font3.colour_index=48#蓝色
style3Row.font = font3 # 设定样式
#style3Row.alignment = al
#统一字体
fontx = xlwt.Font() # 为样式创建字体
fontx.name = u'微软雅黑' 
fontx.height=0x00C8*1.2 # C8 in Hex (in decimal) = 10 points in height.
#灰色背景
styleGrayRow = xlwt.XFStyle() 
styleGrayRow.borders = borders
styleGrayRow.font=fontx
patternGray = xlwt.Pattern() 
patternGray.pattern =xlwt. Pattern.SOLID_PATTERN 
patternGray.pattern_fore_colour = xlwt.Style.colour_map['gray25'] 
styleGrayRow.pattern=patternGray
styleGrayRow.alignment = al

#浅绿背景
styleLightGreenRow = xlwt.XFStyle() 
styleLightGreenRow.borders = borders
styleLightGreenRow.font=fontx  
patternLightGreen = xlwt.Pattern() 
patternLightGreen.pattern =xlwt. Pattern.SOLID_PATTERN 
patternLightGreen.pattern_fore_colour = xlwt.Style.colour_map['light_green'] 
styleLightGreenRow.pattern=patternLightGreen
styleLightGreenRow.alignment=al
styleLightGreenRow.alignment = al
#绿色背景
styleGreenRow = xlwt.XFStyle() 
styleGreenRow.borders = borders
styleGreenRow.font=fontx
patternGreen = xlwt.Pattern() 
patternGreen.pattern =xlwt. Pattern.SOLID_PATTERN 
patternGreen.pattern_fore_colour = xlwt.Style.colour_map['green'] 
styleGreenRow.pattern=patternGreen
#设置单元格背景色
styleGreenRow.pattern = patternGreen
styleGreenRow.alignment = al

#红色字
redstyle = xlwt.XFStyle() 
redstyle.borders = borders
fontred= xlwt.Font() # 为样式创建字体
fontred.name = u'微软雅黑' 
fontred.height=0x00C8*1.2 # C8 in Hex (in decimal) = 10 points in height.
fontred.colour_index=2#红色
redstyle.font=fontred
redstyle.pattern=patternGray
redstyle.alignment = al
#通用style
styleCommon = xlwt.XFStyle() 
styleCommon.borders = borders
styleCommon.font=fontx
styleCommon.alignment = al
#模板内容
worksheet.write_merge(0, 0,0,12, label = "产品开发部工作周报",style=stylefirstRow)
worksheet.write_merge(1, 1,5,7, label = "张三",style=stylesecondRow)
worksheet.write_merge(2, 2,0,12, label = "注：每周五时做本周工作总结（关表），同时做下周工作计划（开表），必要时下周一调整；",style=style3Row)
worksheet.write(3, context, label = "内容",style=styleGrayRow)
worksheet.write(3, ids, label = "序号",style=styleGrayRow)
worksheet.write(3, thing, label = "工作事项",style=styleGrayRow)
worksheet.write(3, project, label = "所属项目",style=styleGrayRow)
worksheet.write(3, skill, label = "技术点分解描述",style=styleGrayRow)
worksheet.write(3, other, label = "需要协调内容（多人接口联调、现场测试、客户试用等等）",style=redstyle)
worksheet.write(3, hour, label = "工作估时",style=styleGrayRow)
worksheet.write(3, planStart, label = "计划开始时间",style=styleGrayRow)
worksheet.write(3, planEnd, label = "计划结束时间",style=styleGrayRow)
worksheet.write(3, start, label = "实际开始时间",style=styleGreenRow)
worksheet.write(3, end, label = "实际开始时间",style=styleGreenRow)
worksheet.write(3, statu, label = "完成状态",style=styleGreenRow)
worksheet.write(3, backup, label = "总结备案",style=styleGreenRow   )          
for line in f:
    line=line.replace('\n','')
    #如果上一次获取到的行为空格，则认为当前行为日期，直接跳过
    if newDay==1:
        newDay=0
    #搜索到##则认为一周结束
    if line.find("====")!=-1:
        break
    else:
        #print strTime
        #如果获取的内容为空格则开始新一天的数据
        if line=="":
            #更新当前的日期
            newDay=1
            now=now+datetime.timedelta(days=-1)
        else:
            newDay=0#置为非新一天的数据
           # print len(line)#==8 or line.count==5)
            if (len(line)==8 or len(line)==5) and line.find(".")!=-1:
                continue#print line.find(".")
            else:
                worksheet.write(i, ids, label = count,style=styleCommon)
                worksheet.write(i, thing, label = line,style=styleCommon)
                worksheet.write(i, project, label = "",style=styleCommon)
                worksheet.write(i, skill, label = "",style=styleCommon)
                worksheet.write(i, other, label = "",style=styleCommon)
                worksheet.write(i, hour, label = "9",style=styleCommon)
                worksheet.write(i, planStart, label = now.strftime("%Y-%m-%d"),style=styleCommon)
                worksheet.write(i, planEnd, label = now.strftime("%Y-%m-%d"),style=styleCommon)
                worksheet.write(i, start, label = now.strftime("%Y-%m-%d"),style=styleCommon)
                worksheet.write(i, end, label = now.strftime("%Y-%m-%d"),style=styleCommon)
                worksheet.write(i, statu, label = "完成",style=styleCommon)
                worksheet.write(i, backup, label = "",style=styleCommon)
                print now.strftime("%Y-%m-%d") +"***"+line
                i=i+1
                count=count+1
    #print strTime
worksheet.write_merge(4, count+4-2,0,0, label = "本周工作总结",style=styleLightGreenRow)
f.close()
workbook.save(r"C:\Users\Dell\Desktop\\"+"产品开发周报_张三_".decode('utf-8').encode('cp936')+datetime.datetime.now().strftime("%Y_%m_%d")+".xlsx")
