import os
import xlwt
import time
import xlrd
import xlwt
import xlutils
from xlutils.copy import copy
class ExcelFileInfo:
    arrsheet = []
    arrfile = []
    arrcolulm = []
    fieldcnt=53
    curcolulm = 1
    def __init__(self, filebasicname):
        self.filebasicname = filebasicname
        for i in range(self.fieldcnt):
            filename=filebasicname+str(i)+".xlsx"
            if os.path.exists(filename):
                os.remove(filename)
            # workfile = xlwt.Workbook(encoding="utf-8")
            # worksheet = workfile.add_sheet("Sheet1",cell_overwrite_ok=True)
            # self.arrsheet.append(worksheet);
            # self.arrfile.append(workfile);
            # self.arrcolulm.append[0]

    def Save(self):
        filename=self.filebasicname+str(0)+".xlsx"
        self.arrfile[0].save(filename)

        filename=self.filebasicname+str(1)+".xlsx"
        self.arrfile[1].save(filename)
        # for i in range(self.fieldcnt):
        #     filename=self.filebasicname+str(i)+".xlsx"
        #     self.arrfile[i].save(filename)
    def AddData(self, strpath):
        self.arrsheet = []
        self.arrfile = []
        for i in range(self.fieldcnt):
            filename=self.filebasicname+str(i)+".xlsx"
            if os.path.exists(filename):
                # 打开已存在的Excel文件
                rd = xlrd.open_workbook(filename)   # 打开文件
                # 复制工作簿到新的变量
                new_workbook = copy(rd)
                sheets = new_workbook.get_sheet(0)   # 读取第一个工作表
                self.arrsheet.append(sheets);
                self.arrfile.append(new_workbook);
            else:
                filename=self.filebasicname+str(i)+".xlsx"
                workfile = xlwt.Workbook(encoding="utf-8")
                worksheet = workfile.add_sheet("Sheet1",cell_overwrite_ok=True)
                self.arrsheet.append(worksheet);
                self.arrfile.append(workfile);


        t = time.time()
        print("strpath=", strpath)
        with open(strpath, 'r') as readfile:
            line = readfile.readline()
            cnt = 1
            idx=0
            lineitr = 0;
            while line:
                # 处理每一行数据
                if cnt>3:
                    line = line.rstrip('\n')
                    # print(line)
                    linedata = line.split(' ')
                    # print(linedata) 
                    # print("len=", len(linedata))
                    for h in range(len(linedata)):
                        # filename=file+str(i)+".xlsx"
                        # print("filename="+filename);
                        # print("idx=", idx, " self.curcolulm=", self.curcolulm)
                        self.arrsheet[h].write(idx,self.curcolulm, linedata[h])
                    idx = idx+1
                cnt = cnt + 1
                line = readfile.readline()
        self.curcolulm = self.curcolulm + 2
        for i in range(self.fieldcnt):
            filename=self.filebasicname+str(i)+".xlsx"
            self.arrfile[i].save(filename)
        # print("name="+strpath+" filebasicname="+self.filebasicname);
        print(f'coast time:{time.time() - t:.8f}s')
        
        # linedata = strline.split(' ')
       
        # # print(linedata) 
        # # self.filebasicname
        # cnt=0
        # for i in range(len(linedata)):
        #     self.arrsheet[i].write(cnt,self.arrcolulm, "内容1")
        #     cnt
            # filename=file+str(i)+".xlsx"
            # print("filename="+filename);


    # def print_car_info(self):
    #     print(f"{self.make} {self.model} {self.year}")

def find_files(directory):
    
    allExcelFileInfo = []
    fileinfo =  ExcelFileInfo("dynamic-pressure-rfile.out");
    # filename="static-pressure-rfile.out1.xlsx"
    # if os.path.exists(filename):
    #     print("exit info")
    # # 打开已存在的Excel文件
    #     rd = xlrd.open_workbook(filename)   # 打开文件
    #         # 复制工作簿到新的变量
    #     new_workbook = copy(rd)
    #     sheets = new_workbook.get_sheet(0)   # 读取第一个工作表
    # return
    allExcelFileInfo.append(fileinfo)
    fileinfo =  ExcelFileInfo("static-pressure-rfile.out");
    allExcelFileInfo.append(fileinfo)
    fileinfo =  ExcelFileInfo("total-pressure-rfile.out");
    allExcelFileInfo.append(fileinfo)
    fileinfo =  ExcelFileInfo("velocity-magnitude-rfile.out");
    allExcelFileInfo.append(fileinfo)
    fileinfo =  ExcelFileInfo("vorticity-magnitude-rfile.out");
    allExcelFileInfo.append(fileinfo)
    
    # fieldcnt=53

    # for i in range(fieldcnt):
    #     # workbook = xlwt.Workbook(encoding="utf-8")
    #     filename="dynamic-pressure-rfile.out"+str(i)+".xlsx"
    #     if os.path.exists(filename):
    #         os.remove(filename)

    # for i in range(fieldcnt):
    #     workfile = xlwt.Workbook(encoding="utf-8")
    #     worksheet = workfile.add_sheet("Sheet1")
    #     arrsheet.append(worksheet);
    #     arrfile.append(workfile);

    for root, dirs, files in os.walk(directory):
        havedata = 0;
        for file in files:
            # print(file)
            # print(os.path.join(root, file))
            
            residx = -1
            for k in range(len(allExcelFileInfo)):
                if allExcelFileInfo[k].filebasicname == file:
                    residx = k
                    break

            if residx == -1:
                continue
            print("residx="+str(residx));
            allExcelFileInfo[residx].AddData(os.path.join(root, file))
            
            # return
    # for k in range(len(allExcelFileInfo)):
    #     allExcelFileInfo[k].Save()
        #     with open(os.path.join(root, file), 'r') as readfile:
        #         line = readfile.readline()
        #         cnt = 1
        #         lineitr = 0;
        #         while line:
        #             # 处理每一行数据
        #             if cnt>3:
        #                 line = line.rstrip('\n')
        #                 print(line)
        #                 linedata = line.split(' ')
        #                 print(linedata) 
        #                 arrsheet = []
        #                 arrfile = []
        #                 for i in range(len(linedata)):
        #                     filename=file+str(i)+".xlsx"
        #                     print("filename="+filename);
        #                     # if os.path.exists(filename):
        #                     #     workbook =open_workbook(filename)
        #                     #     sheet = workbook.sheet_by_index(0) #第一个sheet
        #                     #     arrfile.append(workbook)
        #                     #     arrsheet.append(sheet);
        #                     #     # print("文件存在")
        #                     # else:
        #                     #     book = xlrd.Workbook(encoding="utf-8")
        #                     #     sheet1 = book.add_sheet('Sheet1')
        #                     #     arrfile.append(workbook)
        #                     #     arrsheet.append(sheet);

        #                 # for i in range(len(linedata)):
        #                 #     # filename=file+str(i)+".xlsx"
        #                 #     # if os.path.exists(filename):
        #                 #     #     print("文件存在")
        #                 #     arrsheet[i].write(i,,str)
        #             cnt = cnt + 1
        #             line = readfile.readline()
        #     break
        # if havedata == 1:
        #     break;
 
# 调用函数进行测试
find_files('o\\')