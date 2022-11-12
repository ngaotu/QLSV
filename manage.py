from tkinter import ttk
from tkinter import *
from PIL import ImageTk, Image
import pymysql
import csv
from openpyxl import Workbook
from tkinter import messagebox
SCREEN_WIDTH = 1280
SCREEN_HEIGHT = 720
window = Tk()

class Student:
    def __init__(self,name,age,birth,gender,id_st,mark):
        self.name = name
        self.age = age
        self.birth = birth
        self.gender =gender
        self.id =id_st
        self.mark = mark
class Manage_list:
    def __init__(self,window):
        self.window = window
        self.window.title('Quan ly hoc sinh')
        self.window.geometry(f'{SCREEN_WIDTH}x{SCREEN_HEIGHT}')
        self.window.attributes('-topmost',True)
        self.window.resizable(False,False)
        self.window['bg'] = '#074463'
        subtitle = Label(self.window, text = 'PHẦN MỀM QUẢN LÝ SINH VIÊN', fg = '#FFFF00',font=('Arial,sans-serif',26,'bold'),bg = '#074463')
        subtitle.place(x= SCREEN_WIDTH//2 -250, y= 30)
        self.name = StringVar()
        self.age = StringVar()
        self.birth = StringVar()
        self.gender = StringVar()
        self.id = StringVar()
        self.mark = StringVar()
        self.search = StringVar()
        self.txt_search = StringVar()
        ##### Manage_Frame #######
        Manage_Frame = Frame(self.window,bg = '#1c5775',bd= 2,highlightbackground='black',highlightthickness=3)
        Manage_Frame.place(x = 0,y = 100,width=420,height=SCREEN_HEIGHT-100)
        name = Label(Manage_Frame,text='Họ và tên',font=('Arial,sans-serif',13,'bold'),fg='white',bg ='#1c5775')
        name.grid(row= 0 , column= 0, pady=10,sticky='w', padx= 5)
        input_name = Entry(Manage_Frame,textvariable=self.name,font=('Arial,sans-serif',13,'bold'),fg='black',bg ='white', width=30)
        input_name.grid(row= 0 , column= 1,pady=10 ,sticky='w',padx= 5)


        age = Label(Manage_Frame,text='Tuổi',font=('Arial,sans-serif',13,'bold'),fg='white',bg ='#1c5775')
        age.grid(row= 1, column= 0 , pady= 10,sticky='w',padx= 5)
        input_age = Entry(Manage_Frame,textvariable=self.age,font=('Arial,sans-serif',13,'bold'),fg='black',bg ='white', width=30)
        input_age.grid(row = 1, column= 1,sticky='w',pady=10,padx= 5)



        birth = Label(Manage_Frame,text='Ngày sinh',font=('Arial,sans-serif',13,'bold'),fg='white',bg ='#1c5775')
        birth.grid(row= 2, column= 0 , pady= 10,sticky='w',padx= 5)
        input_birth= Entry(Manage_Frame,textvariable=self.birth,font=('Arial,sans-serif',13,'bold'),fg='black',bg ='white', width=30)
        input_birth.grid(row = 2, column= 1,sticky='w',pady=10,padx= 5)
        

        gender = Label(Manage_Frame,text='Giới tính',font=('Arial,sans-serif',13,'bold'),fg='white',bg ='#1c5775')
        gender.grid(row= 3, column= 0 , pady= 10,sticky='w',padx= 5)
        input_gender= ttk.Combobox(Manage_Frame,textvariable=self.gender,font=('Arial,sans-serif',13,'bold'), width=28, state='readonly')
        input_gender['values'] = ['Nam','Nữ','Giới tính khác']
        input_gender.grid(row = 3, column= 1,sticky='w',pady=10,padx= 5)


        id_st = Label(Manage_Frame,text='Mã sinh viên',font=('Arial,sans-serif',13,'bold'),fg='white',bg ='#1c5775')
        id_st.grid(row= 4, column= 0 , pady= 10,sticky='w',padx= 5)
        input_id= Entry(Manage_Frame,textvariable=self.id,font=('Arial,sans-serif',13,'bold'),fg='black',bg ='white', width=30)
        input_id.grid(row = 4, column= 1,sticky='w',pady=10,padx= 5)

        mark = Label(Manage_Frame,text='GPA',font=('Arial,sans-serif',13,'bold'),fg='white',bg ='#1c5775')
        mark.grid(row= 5, column= 0 , pady= 10,sticky='w',padx= 5)
        input_mark= Entry(Manage_Frame,textvariable=self.mark,font=('Arial,sans-serif',13,'bold'),fg='black',bg ='white', width=30)
        input_mark.grid(row = 5, column= 1,sticky='w',pady=10,padx= 5)


        ############ btn_frame ############
        btn_frame = Frame(Manage_Frame, bg = '#1c5775')
        btn_frame.place(x = 15,y = SCREEN_HEIGHT-220, width= 390, height= 100)
        add_btn = Button(btn_frame,text = 'Thêm',command=self.add_student,bg='#5f9ea0',font=('Arial,sans-serif',13,'bold'),fg = 'white',width=10).grid(row = 0, column= 0, padx=10,pady=10)
        del_btn = Button(btn_frame,text = 'Xóa',bg='#5f9ea0',command=self.del_student,font=('Arial,sans-serif',13,'bold'),fg = 'white',width=10).grid(row = 0, column= 1, padx=10,pady=10)
        update_btn = Button(btn_frame,text = 'Cập nhật',command=self.update_student,bg='#5f9ea0',font=('Arial,sans-serif',13,'bold'),fg = 'white',width=10).grid(row = 0, column= 2, padx=10,pady=10)
        clear_btn = Button(btn_frame,text = 'Làm mới',bg='#5f9ea0',font=('Arial,sans-serif',13,'bold'),fg = 'white',width=10,command= self.clear_console).grid(row = 1, column= 1, padx=10,pady=10)
        close_btn = Button(btn_frame,text = 'Thoát',bg='#5f9ea0',font=('Arial,sans-serif',13,'bold'),fg = 'white',width=10,command= self.close).grid(row = 1, column= 2, padx=10,pady=10)

        #### console_frame ####
        console_frame = Frame(self.window, bg = '#1c5775',highlightbackground='black',highlightthickness=3)
        console_frame.place(x = 460, y = 100, width=SCREEN_WIDTH - 460 , height= SCREEN_HEIGHT -100)
        search_infor = Label(console_frame,text = 'Search by',fg = 'white',bg='#1c5775',  font=('Arial,sans-serif',16,'bold'))
        search_infor.grid(row = 0, column= 0,padx=10,pady= 20)
        search = ttk.Combobox(console_frame,textvariable=self.search,font=('Arial,sans-serif',13,'bold'), width=20, state='readonly')
        search['values'] = ['Name','Age','Date','Gender','ID','GPA']
        search.grid(row = 0, column= 1,padx=10,pady= 20)
        search_inputbtn = Entry(console_frame,textvariable=self.txt_search,fg = 'black',bg='white',width= 20,  font=('Arial,sans-serif',13,'bold'))
        search_inputbtn.grid(row = 0, column= 2,padx=10,pady= 20)
        search_btn = Button(console_frame,text = 'Search',command=self.search_student,fg = 'black',bg='white',width= 10,  font=('Arial,sans-serif',13,'bold'))
        search_btn.grid(row = 0, column= 3,padx=10,pady= 20)
        show_btn = Button(console_frame,text = 'Show',command=self.fetch_data,fg = 'black',bg='white',width= 10,  font=('Arial,sans-serif',13,'bold'))
        show_btn.grid(row = 0, column= 4,padx=10,pady= 20)

        ### Table_frame   ####
        Table_frame = Frame(console_frame, bg = '#1c5775',highlightbackground='white',highlightthickness=3)
        Table_frame.place(x = 17, y = 70, width=SCREEN_WIDTH - 460-40 , height= SCREEN_HEIGHT -200)
        scroll_x = Scrollbar(Table_frame,orient=HORIZONTAL)
        scroll_y = Scrollbar(Table_frame,orient=VERTICAL)
        self.list = ttk.Treeview(Table_frame,columns=('ID','Name','Age','Date','Gender','GPA'),show='headings', xscrollcommand=scroll_x,yscrollcommand=scroll_y)
        scroll_x.pack(side=BOTTOM,fill = X)
        scroll_y.pack(side = RIGHT,fill = Y)
        scroll_x.config(command=self.list.xview)
        scroll_y.config(command=self.list.yview)

        self.list.heading('ID',text = 'Mã sinh viên')
        self.list.heading('Name',text = 'Tên')
        self.list.heading('Age',text = 'Tuổi')
        self.list.heading('Date',text = 'Ngày sinh')
        self.list.heading('Gender',text = 'Giới tính')
        self.list.heading('GPA',text = 'GPA')
        self.list.column('ID', width=50)
        self.list.column('Name', width=100)
        self.list.column('Age', width=20)
        self.list.column('GPA', width=20)
        self.list.column('Date', width=100)
        self.list.column('Gender', width=20)
        self.list.pack(fill=BOTH, expand= True)

    def close(self):
        self.window.destroy()
    def add_student(self):
        try:
            con = pymysql.connect(host='127.0.0.1',user='root',password='ngaotu66',database='student')
            cur = con.cursor()
            #chen gia tri vao  table in my database
            cur.execute("insert into student values (%s, %s, %s, %s,%s,%s)",(self.id.get(),
            self.name.get(),
            self.age.get(),
            self.birth.get(), 
            self.gender.get(),
            self.mark.get()))
            con.commit()
            con.close()
        except:
            messagebox.showerror("Error","Vui long nhap lai")

    def fetch_data(self):
        con = pymysql.connect(host='127.0.0.1',user = 'root',password='ngaotu66',database='student')
        cur = con.cursor()
        cur.execute("select * from student")
        rows = cur.fetchall()
        table_name = []
        for i in cur.description:
            table_name.append(i[0])
        print(table_name)
        if len(rows)!= 0:
            self.list.delete(*self.list.get_children())
            for row in rows:
                self.list.insert('',END,values=row)
            

            wb = Workbook()
            ws = wb.active
            ws.title = "mysql.data"
            ws.append(table_name)
            for row in rows:
                ws.append(row)
            wb.save("student.csv")
            con.commit()
            # print(self.list.get_children())
            
        con.close()
    def del_student(self):
        con = pymysql.connect(host='127.0.0.1', user='root',password='ngaotu66',database='student')
        cur = con.cursor()
        cur.execute("delete from student where ID=%s ",self.id.get())
        con.commit()
        con.close()
        self.clear_console()
    def update_student(self):
        con = pymysql.connect(host = '127.0.0.1',user = 'root',password='ngaotu66',database='student')
        cur = con.cursor()
        cur.execute("update student set Name=%s, Age=%s, Date=%s, Gender=%s, GPA=%s where ID=%s", (
            self.name.get(),
            self.age.get(),
            self.birth.get(), 
            self.gender.get(),
            self.mark.get(),
            self.id.get()))
        con.commit()
        self.clear_console()
        con.close()
    def search_student(self):
        con = pymysql.connect(host='127.0.0.1',user='root',password='ngaotu66',database='student')
        cur = con.cursor()
        cur.execute("select * from student where "+ str(self.search.get()) +" Like '%"+str(self.txt_search.get())+"%'")
        rows = cur.fetchall()
        if len(rows) !=0:
            self.list.delete(*self.list.get_children())
            for row in rows:
                self.list.insert('',END,values=row)
            con.commit()
        con.close()
    def clear_console(self):
        self.name.set("")
        self.age.set("")
        self.birth.set("")
        self.gender.set("")
        self.id.set("")
        self.mark.set("")
obj = Manage_list(window)
window.mainloop()
