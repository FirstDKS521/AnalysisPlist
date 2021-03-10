#!/usr/bin/python3
import os
import zipfile, shutil
from biplist import *
import xlwt

fileType = '.zip'
fileUnzip = 'ipaUnzip'

rootPath = os.getcwd()
unzipPath = rootPath + '/' + fileUnzip

workbook = xlwt.Workbook(encoding="utf-8")
sheet = workbook.add_sheet("pvuv_sheet")

# zip_src: 是zip文件的全路径
# dst_dir: 是要解压到的目的文件夹
def unzip_file(zip_src, dst_dir):
    r = zipfile.is_zipfile(zip_src)
    if r:
        fz = zipfile.ZipFile(zip_src, 'r')
        for file in fz.namelist():
            fz.extract(file, dst_dir)
    else:
        print('This is not zip')

# 在当前路径创建一个新的文件夹，用于存放解压之后的文件
if os.path.exists(unzipPath):
    # 删除原来的文件夹（包括文件夹里面的文件）
    shutil.rmtree(unzipPath)

os.mkdir(fileUnzip)

os.chdir(rootPath + '/ipas')

for ipaName in os.listdir(os.getcwd()):
    portion = os.path.splitext(ipaName)  # 分离文件名与扩展名
    if portion[1] == '.ipa':
        newName = portion[0] + fileType
        os.rename(ipaName, newName)

print('====== 完成更改文件类型 ======')

for ipaZip in os.listdir(os.getcwd()):
    print('开始解压~%s' % ipaZip)
    # 当前zip的文件路径
    zipPath = os.getcwd() + '/' + ipaZip
    # 解压路径
    portion = os.path.splitext(ipaZip)  # 分离文件名与扩展名
    os.mkdir(unzipPath + '/' + portion[0])
    ipaUnzipPath = unzipPath + '/' + portion[0]

    # 开始解压
    unzip_file(zipPath, ipaUnzipPath)

print('====== 完成解压 ======')

for unzipIpaName in os.listdir(unzipPath):
    # 进入Payload文件夹下
    tempPath = unzipPath + '/' + unzipIpaName + '/Payload'
    if os.path.exists(tempPath):
        os.chdir(tempPath)

        for app in os.listdir(os.getcwd()):
            portion = os.path.splitext(app)  # 分离文件名与扩展名
            if portion[1] == '.app':
                newName = portion[0]
                os.rename(app, newName)

# 处理Plist文件
dataList = []
def handlePlist(plist):
    print(plist)

    tempList = []

    appName = plist['CFBundleName']
    if len(appName) == 0:
        appName = plist['CFBundleDisplayName']
    tempList.append(appName) # APP名字
    print(appName)

    schemesList = plist['LSApplicationQueriesSchemes']
    schemesStr = ''
    for scheme in schemesList:
        if len(scheme):
            if len(schemesStr):
                schemesStr = schemesStr + '\n' + scheme
            else:
                schemesStr = scheme

    schemesCount = len(schemesList) # 列表个数

    urlTypeStr = ''
    urlTypesList = plist['CFBundleURLTypes']
    urlTypeStr = 'CFBundleURLSchemes'
    for itemType in urlTypesList:
        if len(itemType) and urlTypeStr in itemType.keys():
            urlsList = itemType[urlTypeStr]
            for urlType in urlsList:
                if len(urlTypeStr):
                    urlTypeStr = urlTypeStr + '\n' + urlType
                else:
                    urlTypeStr = urlType

    tempList.append(schemesCount)
    tempList.append(urlTypeStr)
    tempList.append(schemesStr)

    dataList.append(tempList)

# 解析Plist文件，导出excel表
for unzipIpaName in os.listdir(unzipPath):
    # 进入Payload文件夹下
    tempPath = unzipPath + '/' + unzipIpaName + '/Payload'
    if os.path.exists(tempPath):
        for app in os.listdir(tempPath):
            plistPath = tempPath + '/' + app + '/info.plist'
            if os.path.exists(plistPath):
                plist = readPlist(plistPath)
                handlePlist(plist)

# 先删除旧的Excel表
excelPath = rootPath + '/result.xls'
if os.path.exists(excelPath):
    os.remove(excelPath)

# 标题
titles = ['软件名', '个数', 'URL Type', 'Scheme']
for index, title in enumerate(titles):
    sheet.write(0, index, title)

# 写入每一行
# datas = [['同花顺', '74', 'IHexinFree', 'weixin\nbaidu\nweibo\ndouyin']]
for row, rowData in enumerate(dataList):
    for index, column_data in enumerate(rowData):
        sheet.write(row + 1, index, column_data)

workbook.save(excelPath)

print('====== 导出Excel完成 ======')