import xlwings as xw
from collections import Counter
app = xw.App(visible=False,add_book=False)
app.display_alerts = False
app.screen_updating = False
file_name = [
    "../020#/数据记录_2022_11_17_14_05_05_099.xlsx",
    "../020#/数据记录_2022_11_17_14_15_05_097.xlsx",
    "../020#/数据记录_2022_11_17_14_25_05_097.xlsx",
    "../020#/数据记录_2022_11_17_14_35_05_096.xlsx",
    "../020#/数据记录_2022_11_17_14_45_05_094.xlsx",
    "../020#/数据记录_2022_11_17_14_49_52_423.xlsx"]
wb = app.books.open(file_name[0])
sht = wb.sheets['sheet1']
#把所有的数据都拷贝到第一个表格里面去
for i in range(1,6):
    wb_row = sht.range("A2").expand('table')
    wb_row_count = wb_row.rows.count
    wb2 = app.books.open(file_name[i])
    sht2 = wb2.sheets['sheet1']
    wb2_row = sht2.range("A2").expand('table')
    wb2_row_count = wb2_row.rows.count
    sht.range(f'A{wb_row_count}').expand('table').value = sht2.range(f"A2:S{wb2_row_count}").value


rng = sht.range("G2").expand('table')
row = rng.rows.count
var_temp = sht.range(f"G2:G{row}").value
var_power = sht.range(f"C2:C{row}").value
min_t = -35.0
max_t = 85.0
agv = []
p_t_list =[]
p_t = {}
i = 0
for x in var_temp:
    if(x <= min_t):
        if(x == min_t):
            agv.append(var_power[i])        #这个温度下的数据，都存起来
    else:
        res = Counter(agv)                  #统计列表中的元素个数，返回个字典{“功率”，"次数""}，key = 功率，value = 次数
        cp = max(res,key=res.get)           #找出字典(就是个map)中最大的元素,即最多的那个元素的
        #print("T:",min_t,"P:",cp)
        temp = [min_t,cp]
        p_t_list.append(temp)
        agv.clear()
        min_t = min_t+5
    i+=1
    #超过85度的就不要了
    if(min_t>max_t):
        break
#print(p_t_list)
#打开要保存的模板文件
excelm = app.books.open("../020#/p_t.xlsx")
#从第二行开始保存
excelm.sheets[0].range("A2").expand("table").value = p_t_list
#保存
excelm.save()
#关闭工作薄，一定要关闭，不然会一直占用着
excelm.close()
#关闭excel，一定要关闭
app.quit()