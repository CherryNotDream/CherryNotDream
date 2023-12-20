def create_visual_code_file():
    code = """\
import pyodbc
import os.path 
from tkinter import *
from tkinter import ttk
from tkinter.messagebox import showerror
from PIL import Image,ImageTk

server = '192.168.1.233'
db = 'demo_wibe'
usname = 'admin'
uspsw = '123456'

dbstring = f'DRIVER={{ODBC Driver 17 for SQL Server}};\
             SERVER={server};DATABASE={db};\
             UID={usname};PWD={uspsw}'
where = ''
orderby = 'ORDER BY Name'
query = f'SELECT *,Price*(100-Discount)/100 AS NewPrice FROM Products \
          INNER JOIN Measures ON Measure=IDMeasure \
          INNER JOIN Manufacturers ON Manufacturer=IDManufacturer \
          INNER JOIN Vendors ON Vendor=IDVendor \
          INNER JOIN Categories ON Category=IDCategory'
new_query = f'{query} {where} {orderby}'
record = []
lbl_find_val = ['Найдено товаров','','из','']
try:
    conn = pyodbc.connect(dbstring)
except Exception:
    showerror('Ошибка подключения','Не удалось подключиться к базе данных, проверьте соединение')

def data():
    global record
    record = []
    try:
        cursor = conn.cursor()
        cursor.execute(new_query)
        for row in cursor.fetchall():
            image = row.Picture if row.Picture else 'picture.png'
            if not os.path.exists(row.Picture):
                image = 'picture.png'
            img1 = Image.open(image).resize((45, 45))
            img = ImageTk.PhotoImage(img1)
            tag = 'sale' if row.MaxDiscount>15 else 'blank'
            if row.Discount != 0:
                row.Price=''.join([u'\u0335{}'.format(c) for c in str(row.Price)])
            line = [img,(row.Article,row.Name,row.MeasureName,
                row.Price,row.NewPrice,row.MaxDiscount,
                row.ManufacturerName,row.VendorName,row.CategoryName,
                row.Discount,row.Amount,row.Description),tag]
            record.append(line)
    except Exception:
        showerror('Сбой подключения','Отсутствует подключение к базе данных, проверьте соединение')
    return record

def tree_fill():
    for i in data_tree.get_children():
        data_tree.delete(i)
    for row in data():
        data_tree.insert('',END,open=True,text='',
                         image=row[0],values=row[1],tag=row[2])

def go_status():
    global where
    global lbl_find_val
    global lbl_find
    try:
        cursor = conn.cursor()
        cursor.execute(f'SELECT COUNT(*) FROM Products {where}')
        lbl_find_val[1]=str(cursor.fetchone()[0])
        if lbl_find_val[1]=='0':
            showinfo(title='Информация', message='Товаров не найдено')
        cursor.execute('SELECT COUNT(*) FROM Products ')
        lbl_find_val[3]=str(cursor.fetchone()[0])
        lbl_find.config(text=' '.join(lbl_find_val))
    except Exception:
        showerror(title='Ошибка', message='Нет соединения с базой данных')

def go_sort(event):
    global orderby
    global new_query
    select = cb_sort.get()
    if select == 'По возрастанию':
        orderby = 'ORDER BY Price'
    elif select == 'По убыванию':
        orderby = 'ORDER BY Price DESC'
    else:
        orderby = 'ORDER BY Name'
    new_query = f'{query} {where} {orderby}'
    tree_fill()

def go_filtr(event):
    global orderby
    global where
    global query
    global new_query
    select = cb_filtr.get()
    if select == 'Менее 10%':
        where = 'WHERE MaxDiscount<10'
    elif select == 'От 10 до 15%':
        where = 'WHERE MaxDiscount>=10 and MaxDiscount<15'
    elif select == '15% и более':
        where = 'WHERE MaxDiscount>=15'    
    else:
        where = ''
    new_query = f'{query} {where} {orderby}'
    tree_fill()
    go_status()

app = Tk()
app.geometry('1300x600')
app.title('Мир тканей')
app.minsize(600,300)
#создаем фреймы для комбобоксов, лэйбла и дерева
cb_frame = Frame(app)
lbl_frame = Frame(app)
tree_frame = Frame(app)
#добавляем комбобоксы
lbl_sort = Label(cb_frame,text='Сортировать ')
cb_sort_val = ['Без сортировки','По возрастанию',
               'По убыванию']
cb_sort = ttk.Combobox(cb_frame,values=cb_sort_val,
                       state='readonly')
lbl_filtr = Label(cb_frame,text='Фильтровать ')
cb_filtr_val = ['Без фильтрации','Менее 10%',
               'От 10 до 15%','15% и более']
cb_filtr = ttk.Combobox(cb_frame,values=cb_filtr_val,
                       state='readonly')
# привязка комбобоксов к функциям
cb_sort.bind("<<ComboboxSelected>>", go_sort)
cb_filtr.bind("<<ComboboxSelected>>", go_filtr)
#публикуем комбобоксы таблицей
lbl_sort.grid(column=0,row=0)
cb_sort.grid(column=1,row=0)
lbl_filtr.grid(column=2,row=0)
cb_filtr.grid(column=3,row=0)
#добавляем лэйбл найдено товаров и сразу публикуем
lbl_find = Label(lbl_frame,text=' '.join(lbl_find_val))
lbl_find.pack()
#добавляем дерево
tree = [['#0','Картинка',     'center',50],
        ['#1','Артикул',      'e',     60],
        ['#2','Наименование', 'w',    150],
        ['#3','Ед.изм.',      'w',     50],
        ['#4','Цена',         'e',     70],
        ['#5','Со скидкой',   'e',     70],
        ['#6','Макс.скидка',  'e',     70],
        ['#7','Производитель','e',     80],
        ['#8','Поставщик',    'w',    120],
        ['#9','Категория',    'w',    120],
        ['#10','Скидка',      'e',     50],
        ['#11','Остаток',     'e',     50],
        ['#12','Описание',    'e',    200]]
columns = [k[0] for k in tree]
style = ttk.Style()
style.configure('data.Treeview',rowheight=50)
data_tree = ttk.Treeview(tree_frame,columns=columns[1:],
                         style='data.Treeview')
data_tree.tag_configure('sale',background='#7fff00')
data_tree.tag_configure('blank',background='white')

for k in tree:
    data_tree.column(k[0],width=k[3],anchor=k[2])
    data_tree.heading(k[0],text=k[1])
    
go_status()
tree_fill()

#публикуем дерево
data_tree.pack(fill=BOTH)

#публикуем фреймы
cb_frame.pack(anchor='e',pady=10,padx=20)
lbl_frame.pack(anchor='w',padx=20)
tree_frame.pack(fill=BOTH)

app.mainloop()
"""
    with open('visual_code_project.py', 'w') as file:
        file.write(code)

if __name__ == "__main__":
    create_visual_code_file()
