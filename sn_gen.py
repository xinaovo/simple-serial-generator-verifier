import hmac
import hashlib
import secrets
import string
import colorama
import struct
import csv
from openpyxl import Workbook
import os

SN_LENGTH = 16
SN_SEP = 4
HASHCODE_LENGTH = 4

VERSION = "1.1.0"
SPLITTER = "========================================"
FAIL_ASCII_ART = """
███████  █████  ██ ██      ███████ ██████  
██      ██   ██ ██ ██      ██      ██   ██ 
█████   ███████ ██ ██      █████   ██   ██ 
██      ██   ██ ██ ██      ██      ██   ██ 
██      ██   ██ ██ ███████ ███████ ██████  
"""
OK_ASCII_ART = """
 ██████  ██   ██ 
██    ██ ██  ██  
██    ██ █████   
██    ██ ██  ██  
 ██████  ██   ██ 
"""

def getHMACSHA256(key, message):
    return hmac.new(key.encode("utf-8"), message.encode("utf-8"), hashlib.sha256).digest()

def concatFullSN(sn: list, hashCode: str):
    snString = '-'.join(sn)
    return snString + '-' + hashCode

def parseFullSN(fullSN: str):
    splitSN = fullSN.split('-')
    snHashCode = splitSN[-1]
    snList = splitSN[:-1]
    return snList, snHashCode

def genSN(length : int, sep: int):
    ALPHABET = string.ascii_uppercase + string.digits
    resultString = ''
    resultArray = []
    for i in range(1, length + 1):
        resultString += secrets.choice(ALPHABET)
        if(len(resultString) == sep or i == length):
            resultArray.append(resultString)
            resultString = ''
    return resultArray

def getHashCode(sn : list, length: int):
    CHARSET = "23456789BCDFGHJKMNPQRTVWXY"
    KEY = "CGSRT179SQUY1ZV4Q7LTQHY3K8M6N3PTQUXQGPW7HTD21N5AIDY93M3193H06HI9OC6K35KPOQNVPXZIBXPY3FZEBZHOXMZC9NHEY5MH9LEIQS7QRUZJUHAB1BWR76PE"
    snString = ''.join(sn)
    hmac = getHMACSHA256(KEY, snString)
    start = ord(hmac[19:20]) & 0x0F
    codeint = struct.unpack(">I", hmac[start : start + 4])[0] & 0x7FFFFFFF

    code = []
    for _ in range(length):
        codeint, i = divmod(codeint, len(CHARSET))
        code.append(CHARSET[i])

    return "".join(code)

def genFullSN(snLength : int, sep: int, hashCodeLength):
    sn = genSN(snLength, sep)
    hashCode = getHashCode(sn, hashCodeLength)
    return concatFullSN(sn, hashCode)

def verifyFullSN(fullSN: str):
    snList, snHashCode = parseFullSN(fullSN)
    hashCode = getHashCode(snList, len(snHashCode))
    return hashCode == snHashCode

def mainMenu():
    print(SPLITTER)
    print("主界面选项:")
    print("1. 生成序列号")
    print("2. 验证序列号")
    print("3. 退出")
    print(SPLITTER)
    choice = str(input("请输入您的选择(数字1-3):"))
    while(choice != "1" and choice != "2" and choice != "3"):
        choice = str(input(colorama.Fore.RED + "无效输入! 请重新输入(数字1-3):" + colorama.Style.RESET_ALL))
    match choice:
        case "1":
            subMenuGenSN()
        case "2":
            subMenuVerifySN()
        case "3":
            os._exit(0)
    mainMenu()

def subMenuGenSN():
    print(SPLITTER)

    num = 0
    while(num < 1):
        try:
            num = int(input("请输入生成个数(正整数):"))
        except:
            print(colorama.Fore.RED + "无效输入! 请输入一个正整数! " + colorama.Style.RESET_ALL)

    genSNList = ["序列号"]
    for i in range(num):
        fullSN = genFullSN(SN_LENGTH, SN_SEP, HASHCODE_LENGTH)
        print(fullSN)
        genSNList.append(fullSN)

    choice = str(input("是否将结果导出到.csv(C)或Excel .xlsx(X)表格文件?(C/X/N):"))
    while(choice != "C" and choice != "c" and choice != "X" and choice != "x" and choice != "N" and choice != "n"):
        choice = str(input(colorama.Fore.RED + "无效输入! 请重新输入(字母C、X或N, C代表导出.csv文件, X代表导出Excel .xlsx文件, N代表\"否\"):" + colorama.Style.RESET_ALL))
    if(choice == "C" or choice == "c"):
        workDir = os.getcwd()
        outputFileAbsolutePath = workDir + "\\序列号生成输出-{}.csv".format(''.join(genSN(4, 5)))
        with open(outputFileAbsolutePath, mode='w', newline='', encoding='utf-8') as file:
            csv_writer = csv.writer(file)
            for row in genSNList:
                csv_writer.writerow([row, " "])
        print(colorama.Fore.GREEN + "结果已保存至{}".format(outputFileAbsolutePath) + colorama.Style.RESET_ALL) 
    elif(choice == "X" or choice == "x"):
        workDir = os.getcwd()
        outputFileAbsolutePath = workDir + "\\序列号生成输出-{}.xlsx".format(''.join(genSN(4, 5)))
            
        wb = Workbook()
        ws = wb.active
        ws.title = "序列号"
        for idx, row in enumerate(genSNList, 1):
            ws.cell(row=idx, column=1, value=row)
            wb.save(outputFileAbsolutePath)
        print(colorama.Fore.GREEN + "结果已保存至{}".format(outputFileAbsolutePath) + colorama.Style.RESET_ALL)

    print(SPLITTER)

def subMenuVerifySN():
    print(SPLITTER)
    sn = str(input("请输入需要验证的序列号(XXXX-XXXX-XXXX-XXXX):"))
    if(verifyFullSN(sn)):
        print(colorama.Fore.GREEN + OK_ASCII_ART + "验证成功! 序列号{}有效!".format(sn) + colorama.Style.RESET_ALL)
    else:
        print(colorama.Fore.RED + FAIL_ASCII_ART + "验证失败! 序列号{}无效!".format(sn) + colorama.Style.RESET_ALL)
    print(SPLITTER)

# Program Entry
if __name__ == "__main__":
    LICENSE_INFO = """本软件以MIT协议分发, 若您继续使用该软件即视为您知晓并同意遵守MIT协议。MIT协议如下所示: 
Copyright (C) 2024 Xina.
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE."""

    print(SPLITTER)
    print("欢迎使用序列号生成器!")
    print("版本：{}".format(VERSION))
    print(LICENSE_INFO)
    print(SPLITTER)
    mainMenu()
