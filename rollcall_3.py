import os, xlrd
#--------------functions----------------
def data_import(filename):
    data = xlrd.open_workbook(filename)
    table = data.sheet_by_index(0)
    nrows_num, ncols_num = table.nrows, table.ncols
    res = [[None]*ncols_num for i in range(nrows_num)] #build pre-conditioned 2D list

    for nrows in range(nrows_num):
        for ncols in range(ncols_num):
        
            cell_value = table.cell(nrows,ncols).value
            if cell_value=='':
                    cell_value='_'
                    res[nrows][ncols] = cell_value
            else: 
                res[nrows][ncols] = cell_value

    return res, nrows_num, ncols_num
def intersection(a,b): #定義取交集的函數
    return list(set(a) & set(b))

def show(list1): #加上頓號
    ss = '、'.join(map(str,list1))
    return ss

def nam2num(list):
    newlist=[]
    for item in list:
        newlist.append(convert[item]) #將人名換成號碼
        newlist = sorted(newlist, key = int) #照號碼排序
    return newlist

def count3(list):
    newlist=[]
    for item in list:
        newlist.append(convert[item]) #將人名換成號碼
        newlist = sorted(newlist, key = int) #照號碼排序
    ss = '、'.join(newlist)
    return ss    


def num2nam(list):#將號碼換成人名
    newlist=[]
    for item in list:
        newlist.append(revert[item]) #將號碼換成人名
        #newlist = sorted(newlist, key = int) #照號碼排序
    #s2 = '、'.join(map(str,newlist))
    return newlist


#----------------main------------------

cwd = os.getcwd()
filenames = os.listdir(cwd)
print('----------',cwd,'--------------')
print('\n'.join(filenames))
os.path.exists('106.04役男輪休表(總表).xlsx') #做一個可以偵測檔名的
#做一個可以browse，可以預設抓取的檔案，也可以自由選取檔案 
print('------------------------------')
targetfile = '106.04役男輪休表(總表).xlsx'
mm = targetfile.split('.')[1][0:2]

print('目前選擇操作的檔案為：' + targetfile )
ask = input("確認是否繼續操作？(y/n)" )
if ask in ['y', 'Y','']:
    pass
else:
    targetfile = input('請輸入要修改的檔名: ')

date_start,date_end = 0,30
print('目前操作範圍為 {m}/{d1}~{m}/{d2}'.format(m=mm,d1=date_start+1,d2=date_end))
ask2 = input("確認是否繼續操作？(y/n)" )
if ask2 in ['y', 'Y','']:
    pass
else:
    date_start = int(input('請輸入開始日: '))
    date_end = int(input('請輸入結束日: ')) 
    pass
#Start Processing
#檔案讀入
res,a1,a2 = data_import(targetfile)

status_table=[]
for rows in res:
    if not isinstance(rows[1],float): #尋找號碼 去除垃圾
        pass 
    else: status_table.append(rows)

number = [int(status_table[num][1]) for num in range(len(status_table))] #把號碼建好 int 

date=[] #建立日期
for i in range(len(res[3])):
    if not isinstance(res[3][i],float):
        pass
    else: 
        date.append(mm+'/'+str(int(res[3][i])))    

names = [status_table[i][2] for i in range(len(status_table))] #建立人名 str
convert = dict(zip(names,number)) #將人名換成號碼
revert = dict(zip(number,names)) #將號碼換成人名

#class soilder(obj):
 #   def __init__(self,name,num):
  #      self.name = name
   #     self.number = num

off = ['○','◎','●'] #定義收假放假
on = '_'
offset = 4 #[4]為4/1的狀態

f = open(mm+'月每日役男狀態.txt','wt',encoding = 'utf8')
print('start running')
for x in range(offset+date_start,offset+date_end):
    d1,d2 = 0, 0
    n1,n2 = 0, 0
    n1_names = []
    n2_names = []
    d1_names = []
    d2_names = []
    off1_names = []
    off2_names = []
    off1, off2 = 0,0
#main loop
    for i in range(len(status_table)): #x為日期
        if (status_table[i][x]==on):
            d1_names.append(status_table[i][2])
        elif (status_table[i][x] in off):
            off1_names.append(status_table[i][2])
    for i in range(len(status_table)): #第二天
        if (status_table[i][x+1]==on):
            d2_names.append(status_table[i][2])
        elif (status_table[i][x+1] in off):
            off2_names.append(status_table[i][2])
            
    leave18_names = intersection(off2_names,d1_names)
    back21_names = intersection(off1_names,d2_names)
    off1_numbers = [str(convert[item]) for item in off1_names]
    
    n1_names = list(d1_names)
    for item in leave18_names:
        n1_names.remove(item)
        
    n1_names += back21_names
    d1,d2,n1 = map(len,[d1_names,d2_names,n1_names])
    
    #if not os.path.exists(mm+'月每日役男狀態.txt'):
    #Output 
    print(mm+'/'+str(x-offset+1),'日間機動警力 :',d1,'夜間機動警力:',n1,'\n' ,file=f)
    print(' 早點名:',show(num2nam(nam2num(d1_names))),'\n',file=f) #bug
    print(' 晚點名:',show(num2nam(nam2num(n1_names))),'\n',file=f)
    print(' 18退勤:',show(leave18_names), '共',len(leave18_names),'人',file=f) #bug
    print(' 21收假:',show(back21_names),'共',len(back21_names),'人',file=f)
    print('役男輪休: ',show(nam2num(off1_names)),'共',len(off1_names),'人',file=f)
    print('(' + show(nam2num(leave18_names)) + '於18放假)',file=f)  
    print('(' + show(nam2num(back21_names)) + '於21收假)',file=f)  
    print('-----------------------------------',file=f)

f.close()
print(mm+'月每日役男狀態.txt 寫入完成')