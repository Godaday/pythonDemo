#pip3 install xlrd 
#pip3 install xlsxwriter
import os,xlrd
import xlsxwriter
import datetime,time
Target_Folder_Path="mergeData/" #合并Excel目标文件夹默认程序当前目录下
Sheet_Index=0#数据在Excel的第几个Sheet中，默认第一个0
First_Data_row=5#数据开始的第一行（排除表头）
mergeDoneData=[]#存储读取数据
def MergeFiles(folderPath):
    for root ,dirs,files  in os.walk(folderPath):
        for name in files:
            fileFullPath =os.path.join(root,name)
            finalData=ReadExcelRow(fileFullPath)
        WriteMargeDataInNewFile(finalData) 
#读取文件内容
def ReadExcelRow(filePath):
   filedata =xlrd.open_workbook(filePath)
   sheets = filedata.sheets()
   if len(sheets)>0:
        print("reading file- :",filePath)
        tragetSheet =sheets[Sheet_Index]
        rowCount=0
        for rownum  in range(tragetSheet.nrows):
            if rownum<First_Data_row-1:
                continue
            currentRow=tragetSheet.row_values(rownum)
            if len(currentRow[1].strip())!=0 and len(currentRow[2].strip())!=0:
                currentRow.insert(0,filePath.replace(Target_Folder_Path,""))#增加一列文件名
                mergeDoneData.append(currentRow)
                rowCount+=1
        print("Rows Count:",rowCount)
   else:
        print(filePath,"has 0 sheet")
   return mergeDoneData
def WriteMargeDataInNewFile(mergeDoneData):
    newfileName= "合并结果"+time.strftime("%Y-%m-%d-%H-%M-%S",time.localtime(time.time()))+".xlsx"
    merageDoneFile=xlsxwriter.Workbook(newfileName)
    merageDoneFile_Sheet = merageDoneFile.add_worksheet()
    for row in range(len(mergeDoneData)):
        for cell in range(len(mergeDoneData[row])):
            value=mergeDoneData[row][cell]
            merageDoneFile_Sheet.write(row,cell,value)
    merageDoneFile.close()
    print('merage excel is  done file name is',newfileName)
    os.system("start explorer %s" % os.getcwd())
if __name__=='__main__':
    MergeFiles(Target_Folder_Path)
