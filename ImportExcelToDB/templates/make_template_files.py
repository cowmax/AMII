import io;
import os;
import random;
import time;
from datetime import datetime;
import threading;

TemplateFileNames = ["DispatchProductOrder",
"DispatchProductOrderDetail",
"Product",
"SalesOrder",
"SalesOrderDetail",
"SalesOrderDetail_PlatformProduc",
"SalesOrderDetail_Product",
"SalesOrder_Invoice",
"SalesOrder_Log",
"SalesOrder_Payment",
"SalesOrder_Sub",
"Store",
"Users"];




def fns(nmb, w):
    s = str(nmb);
    l = len(s);
    d = w - l;
    if  d > 0 :
        for i in range(0,d):
            s = '0' + s;

    return s;
    pass;

def getTimeStamp():
    dt = datetime.now()
    ts = "";
    ts = str(dt.year) + fns(str(dt.month),2) + fns(str(dt.day),2) + fns(str(dt.hour),2) + fns(str(dt.minute),2) + fns(str(dt.second),2) + fns(str(dt.microsecond),2);
    return ts;
    pass;

def getverNum():
    pass;


class TmplInfo:
    pass;

def thread_main(ti):
    stm = random.randint(5, 30);
    time.sleep(stm); # 让线程等待片刻
    ti.id = getTimeStamp();

    # 获得线程名
    thrdName = threading.currentThread().getName()
    # 创建文件，并写入少许数据
    fnm = ti.name+"."+ti.id+"."+ hex(ti.version)[2:]+".xlsx";
    tmplFile = open(fnm, "w+", encoding="utf-8");
    tmplFile.write(thrdName + "," + fnm);
    tmplFile.close();

    print (thrdName, ":", ti.name, ti.id, ti.version);

    pass; # END 线程函数

# ------------- MAIN -------------
global count, mutex
wrkThreads = []

fileCount = len(TemplateFileNames);
for i in range(1, fileCount*5):
    idxFnm = random.randint(0, fileCount -1);
    verNum = random.randint(1, 0xffff);
    # 准备文件名称
    ti = TmplInfo();
    ti.name = TemplateFileNames[idxFnm];
    ti.version = verNum;

    # 先创建线程对象
    wrkThreads.append(threading.Thread(target=thread_main, args=(ti,)));

# 启动所有线程
for t in wrkThreads:
    t.start()

# 主线程中等待所有子线程退出
for t in wrkThreads:
    t.join()  
     