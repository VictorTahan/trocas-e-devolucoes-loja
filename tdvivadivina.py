from tkinter import *
from tkinter import ttk,messagebox
from tkcalendar import DateEntry
import pyodbc
import datetime
root = Tk()
connection_data = (
    "DRIVER={SQL Server};"
    "SERVER=DESKTOP-5FNV2D8;"
    "DATABASE=TROCAS_DEVOLUCOES;"    
)
class Funcs():
    def db_connect(self):
        try:
            self.conn = pyodbc.connect(connection_data)
            self.cursor = self.conn.cursor()
        except pyodbc.Error as e1:
            messagebox.showerror('Conncetion failed! Error:',f'{e1}')
            return
    def db_disconnect(self):
        try:
            self.conn.close()
        except pyodbc.Error as e2:
            messagebox.showerror('Disconnection failed! Error:',f'{e2}')
            return
    def limpar_tela(self):
        self.data_tb.set_date(datetime.date.today())
        self.cliente_entry.delete(0,END)
        self.motivo_combobox.set("")
        self.obs_entry.delete(0,END)
        self.vendedora_entry.delete(0,END)
        self.frete_combobox.set("")
        self.status_combobox.set("")
    def variaves(self):
        self.data = self.data_tb.get()
        self.cliente = self.cliente_entry.get()
        self.motivo = self.motivo_combobox.get()
        self.obs = self.obs_entry.get()
        self.vendedora = self.vendedora_entry.get()
        self.frete = self.frete_combobox.get()
        self.status = self.status_combobox.get()
        self.procurar = self.procurar_entry.get()
    def treeselection(self):
        self.tree.delete(*self.tree.get_children())
        self.db_connect()
        linha = self.cursor.execute(f"""SELECT * FROM TROCAS_DEVOLUCOES ORDER BY CD_TROCA""")
        for i in linha:
            self.tree.insert("",END,values=list(i))
        self.db_disconnect()
    def registrar(self):
        self.variaves()
        self.db_connect()
        try:
            self.cursor.execute(f"""INSERT INTO TROCAS_DEVOLUCOES (DS_DATA,NM_CLIENTE,DS_MOTIVO,DS_OBS
                                ,NM_VENDEDORA,DS_FRETE,DS_STATUS) VALUES(?,?,?,?,?,?,?)""",
                                 (self.data,self.cliente,self.motivo,self.obs,self.vendedora,self.frete,self.status) )
            self.conn.commit()
        except pyodbc.Error as e3:
            print(f'Erro: {e3}')
        finally:
            self.db_disconnect()
            self.treeselection()
            self.limpar_tela()
    def OnDoubleClick(self,event):
        self.limpar_tela()
        self.tree.selection()
        for n in self.tree.selection():
            c1,c2,c3,c4,c5,c6,c7,c8 = self.tree.item(n,'values')
            self.selected_code = c1
            try:
                date_object = datetime.datetime.strptime(c2, "%Y-%m-%d").date()
            except ValueError:
                messagebox.showerror("Data inválida", f"A data {c2} não está no formato esperado.")
                return
            self.data_tb.set_date(date_object)
            self.cliente_entry.insert(END,c3)
            self.motivo_combobox.set(c4)
            self.obs_entry.insert(END,c5)
            self.vendedora_entry.insert(END,c6)
            self.frete_combobox.set(c7)
            self.status_combobox.set(c8)
            self.editar_button = Button(self.frame1,text='Editar',bg='#006057',fg='white',bd=2,font=('verdana',8,'bold'),command=self.editar)
            self.editar_button.place(relx=0.7,rely=0.9)
            self.deletar_button = Button(self.frame1,text='Deletar',bg='#006057',fg='white',bd=2,font=('verdana',8,'bold'),command=self.deletar)
            self.deletar_button.place(relx=0.8,rely=0.9)
    def deletar(self):
        self.variaves()
        self.db_connect()
        self.cursor.execute(f"""DELETE FROM TROCAS_DEVOLUCOES WHERE CD_TROCA =?""",(self.selected_code))
        self.conn.commit()
        self.db_disconnect()
        self.limpar_tela()
        self.treeselection()
        self.editar_button.place_forget()
        self.deletar_button.place_forget()
    def editar(self):
        self.variaves()
        self.db_connect()
        self.cursor.execute(f"""UPDATE TROCAS_DEVOLUCOES SET DS_DATA=?,NM_CLIENTE=?,DS_MOTIVO=?,DS_OBS=?,
                            NM_VENDEDORA=?,DS_FRETE=?,DS_STATUS=? WHERE CD_TROCA=?""",
                            (self.data,self.cliente,self.motivo,self.obs,self.vendedora,self.frete,self.status,
                            self.selected_code))
        self.conn.commit()
        self.db_disconnect()
        self.limpar_tela()
        self.treeselection()
        self.editar_button.place_forget()
        self.deletar_button.place_forget()
    def procurar_cliente(self):
        self.variaves()
        search = self.procurar.strip()
        self.tree.delete(*self.tree.get_children())
        self.db_connect()
        query = f"""SELECT * FROM TROCAS_DEVOLUCOES WHERE NM_CLIENTE LIKE ? ORDER BY CD_TROCA"""
        self.cursor.execute(query,(f'%{search}%',))
        for i in self.cursor:
            self.tree.insert("", END, values=list(i))
        self.db_disconnect()
class App(Funcs):
    def __init__(self):
        self.root = root
        self.tela()
        self.frames()
        self.treeview()
        self.treeselection()
        self.widgets()
        root.mainloop()
    def tela(self):
        self.root.title("Trocas e Devoluções")
        self.root.configure(background='#15a79c')
        self.root.geometry('1000x600')
        self.root.resizable(False,False) 
    def frames(self):
        self.frame1 = Frame(self.root,bd=4,bg="#A0F4ED",highlightbackground='#DCFFFC',highlightthickness=3)
        self.frame1.place(relx=0.02,rely=0.02,relwidth=0.96,relheight=0.46)
        self.frame2 = Frame(self.root,bd=4,bg="#A0F4ED",highlightbackground='#DCFFFC',highlightthickness=3)
        self.frame2.place(relx=0.02,rely=0.5,relwidth=0.96,relheight=0.46)
    def widgets(self):
        self.data_label = Label(self.frame1,text='Data:',font=('verdana',8,'bold'),bg="#A0F4ED")
        self.data_label.place(relx=0.1,rely=0.1)
        self.data_tb = DateEntry(self.frame1,date_pattern='dd/mm/yyyy',showweeknumbers=False,background='#15a79c',state='readonly')
        self.data_tb.place(relx=0.2,rely=0.1)
        self.cliente_label = Label(self.frame1,text='Cliente:',font=('verdana',8,'bold'),bg="#A0F4ED")
        self.cliente_label.place(relx=0.1,rely=0.3)
        self.cliente_entry = Entry(self.frame1)
        self.cliente_entry.place(relx=0.22,rely=0.3,relwidth=0.63)
        self.motivo_label = Label(self.frame1,text='Motivo:',font=('verdana',8,'bold'),bg="#A0F4ED")
        self.motivo_label.place(relx=0.4,rely=0.1)
        motivos = ['Peça não serviu','Cliente não gostou do(s) produto(s)','Produto(s) enviado(s) com defeito','Produto(s) enviado(s) errado','Outro(s)']
        self.motivo_combobox = ttk.Combobox(self.frame1,values=motivos,state='readonly')
        self.motivo_combobox.place(relx=0.5,rely=0.1,relwidth=0.35)
        self.obs_label = Label(self.frame1,text='Observação:',font=('verdana',8,'bold'),bg="#A0F4ED")
        self.obs_label.place(relx=0.1,rely=0.5)
        self.obs_entry = Entry(self.frame1)
        self.obs_entry.place(relx=0.22,rely=0.5,relwidth=0.63)
        self.vendedora_label = Label(self.frame1,text='Vendedora:',font=('verdana',8,'bold'),bg="#A0F4ED")
        self.vendedora_label.place(relx=0.1,rely=0.7)
        self.vendedora_entry = Entry(self.frame1)
        self.vendedora_entry.place(relx=0.22,rely=0.7,relwidth=0.25)
        self.frete_label = Label(self.frame1,text='Frete:',font=('verdana',8,'bold'),bg="#A0F4ED")
        self.frete_label.place(relx=0.5,rely=0.7)
        fretes = ['Pago pelo cliente','Etiqueta Reversa']
        self.frete_combobox = ttk.Combobox(self.frame1,values=fretes,state='readonly')
        self.frete_combobox.place(relx=0.6,rely=0.7)
        self.status_label = Label(self.frame1,text='Status:',font=('verdana',8,'bold'),bg="#A0F4ED")
        self.status_label.place(relx=0.1,rely=0.9)
        status = ['Pendente','Resolvido']
        self.status_combobox = ttk.Combobox(self.frame1,values=status,state='readonly')
        self.status_combobox.place(relx=0.22,rely=0.9)
        self.registrar_button = Button(self.frame1,text='Registrar',bg='#006057',fg='white',bd=2,font=('verdana',8,'bold'),command=self.registrar)
        self.registrar_button.place(relx=0.9,rely=0.9)
        self.procurar_label = Label(self.frame2,text='Procurar:',font=('verdana',8,'bold'),bg="#A0F4ED")
        self.procurar_label.place(relx=0.01,rely=0.9)
        self.procurar_entry = Entry(self.frame2)
        self.procurar_entry.place(relx=0.1,rely=0.9,relwidth=0.2)
        self.procurar_button = Button(self.frame2,text='Procurar',bg='#006057',fg='white',bd=2,font=('verdana',8,'bold'),command=self.procurar_cliente)
        self.procurar_button.place(relx=0.32,rely=0.9)
    def treeview(self):
        self.tree = ttk.Treeview(self.frame2,height=3,column=('c1','c2','c3','c4','c5','c6','c7','c8'))
        self.tree.heading('#0',text='')
        self.tree.heading('c1',text='Cód')
        self.tree.heading('c2',text='Data')
        self.tree.heading('c3',text='Cliente')
        self.tree.heading('c4',text='Motivo')
        self.tree.heading('c5',text='Observação')
        self.tree.heading('c6',text='Vendedora')
        self.tree.heading('c7',text='Frete')
        self.tree.heading('c8',text='Status')

        self.tree.column('#0',width=0,stretch=NO)
        self.tree.column('c1',width=10)
        self.tree.column('c2',width=40)
        self.tree.column('c3',width=100)
        self.tree.column('c4',width=115)
        self.tree.column('c5',width=115)
        self.tree.column('c6',width=40)
        self.tree.column('c7',width=40)
        self.tree.column('c8',width=40)

        self.tree.place(relx=0.01,rely=0.01,relwidth=0.96,relheight=0.8)
        
        self.scroll_lista1 = Scrollbar(self.frame2,orient='vertical')
        self.tree.configure(yscroll=self.scroll_lista1.set)
        self.scroll_lista1.place(relx=0.97,rely=0.01,relwidth=0.02,relheight=0.8)
        
        self.scroll_lista2 = Scrollbar(self.frame2,orient='horizontal')
        self.tree.configure(xscroll=self.scroll_lista2.set)
        self.scroll_lista2.place(relx=0.01,rely=0.8,relwidth=0.98,relheight=0.05)
        self.tree.bind("<Double-1>",self.OnDoubleClick)
App()