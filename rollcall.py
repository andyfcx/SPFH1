import os
cwd = os.getcwd()
filenames = os.listdir(cwd)
print('------------------------------')
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

#檔案讀入
import xlrd
#from collections import Counter

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

res,a1,a2 = data_import(targetfile)

status_table=[]
for rows in res:
    if not isinstance(rows[1],float): #尋找號碼 去除垃圾
        pass 
    else: status_table.append(rows)
number = [int(status_table[num][1]) for num in range(len(status_table))] #把號碼建好

date=[] #建立日期
for i in range(len(res[3])):
    if not isinstance(res[3][i],float):
        pass
    else: 
        date.append(mm+'/'+str(int(res[3][i])))    

names=[status_table[i][2] for i in range(len(status_table))] #建立人名
convert = dict(zip(names,number)) #將人名換成號碼


c1,c2='○','◎' #定義收假放假
off = [c1 , c2,'●']
on = '_'
offset = 4 #[4]為4/1的狀態
def intersection(a,b): #定義取交集的函數
    return list(set(a) & set(b))
def show(list1): #加上頓號
    map(str,list1)
    ss = '、'.join(list1)
    return ss
def count3(list):#加上頓號並排序
    newlist=[]
    for item in list:
        newlist.append(convert[item])
        newlist = sorted(newlist, key = int)
    s2 = '、'.join(map(str,newlist))
    return s2

f=open(mm+'月每日役男狀態.txt','rt')
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
  
    print(mm+'/'+str(x-offset+1),'日間機動警力 :',d1,'夜間機動警力:',n1,'\n')
    print(' 早點名:',show(d1_names),'\n')
    print(' 晚點名:',show(n1_names),'\n')
    print(' 18退勤:',show(leave18_names), '共',len(leave18_names),'人\n')
    print(' 21收假:',show(back21_names),'共',len(back21_names),'人\n')
    print('役男輪休: \n',count3(off1_names),'共',len(off1_names),'人\n')
    print('(' + count3(leave18_names) + '於18放假)')  
    print('(' + count3(back21_names) + '於21收假)')  
    print('-----------------------------------')
f.close()



