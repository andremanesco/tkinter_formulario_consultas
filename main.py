## Formulário de Registro de Consultas de itens de uma empresa fictícia, visando a criação de um processo padronizado
## de coleta e armazenamento de dados para possíveis análises futuras de mercado.

import tkinter
from tkinter import *
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
from datetime import date
import re
import getpass
import webbrowser
import pyodbc


class Register:
    def __init__(self, root):
        self.root = root
        self.root.title("Registration Window")
        self.root.geometry("1350x700+0+0")
        # self.root.config(bg="white")
        # self.root.wm_attributes("-transparentcolor", "blue")

        # - Background Image
        img = Image.open("D:/Projetos/formulario_consultas/img/bg.png")
        image = img.resize((1350, 700))
        self.bg = ImageTk.PhotoImage(image)
        bg = Label(self.root, image=self.bg).place(x=0, y=0)

        # - Data Atual
        self.infoDataHoje = date.today().strftime("%d/%m/%Y")
        dataHoje = Label(self.root,
                         text="Data: " + self.infoDataHoje,
                         font=("times new roman", 15, "bold"),
                         bg="#EEF1F2",
                         fg="black")
        dataHoje.place(x=1160, y=100)

        # - Usuário
        self.infoUser = getpass.getuser()

        # - Itens consultados
        txtConsulta = Label(self.root,
                            text="Itens Consultados",
                            font=("times new roman", 20, "bold"),
                            bg="#EEF1F2",
                            anchor="w")
        txtConsulta.place(x=788, y=204, width=214, height=24)

        # - Name
        ## [TEXTO] Nome
        textName = ttk.Label(self.root,
                             text="Nome*",
                             font=("times new roman", 14, "bold"),
                             background="#EEF1F2",
                             anchor="w")
        textName.place(x=331, y=185, width=80, height=25)

        ## [INPUT] Nome
        self.txt_Name = Entry(self.root,
                              font=("times new roman", 15),
                              bg="lightgray")
        self.txt_Name.place(x=331, y=210, width=300)

        # - Email
        ## [TEXTO] Email
        textEmail = Label(self.root,
                          text="E-mail*",
                          font=("times new roman", 14, "bold"),
                          bg="#EEF1F2",
                          anchor="w",
                          fg="black")
        textEmail.place(x=331, y=255, width=80, height=25)

        boxEmailValidationTxt = Label(self.root,
                                      text="",
                                      font=("times new roman", 11, "italic"),
                                      bg="#EEF1F2",
                                      fg="Red")
        boxEmailValidationTxt.place(x=480, y=257, width=100, height=25)

        ## [INPUT] Email
        self.txt_Email = Entry(self.root,
                               font=("times new roman", 15),
                               bg="lightgray")
        regEmail = self.root.register(self.checkEmail)
        self.wdgLst = boxEmailValidationTxt
        self.txt_Email.config(validate="focusout", validatecommand=(regEmail, '%P'))
        self.txt_Email.place(x=331, y=280, width=300)

        # - Telefone
        ## [TEXTO] Telefone
        textTelefone = Label(self.root,
                             text="Telefone",
                             font=("times new roman", 14, "bold"),
                             bg="#EEF1F2",
                             anchor="w",
                             fg="black")
        textTelefone.place(x=331, y=325, width=80, height=25)

        ## [INPUT] Telefone
        self.txt_Telefone = Entry(self.root,
                                  font=("times new roman", 15),
                                  bg="lightgray")
        self.txt_Telefone.place(x=331, y=350, width=300)

        # - Estado
        ## [TEXTO] Estado
        textEstado = Label(self.root,
                           text="Estado*",
                           font=("times new roman", 14, "bold"),
                           bg="#EEF1F2",
                           anchor="w",
                           fg="black")
        textEstado.place(x=331, y=395, width=80, height=25)

        ## [INPUT] Estado
        self.cmb_Estado = ttk.Combobox(self.root,
                                       font=("times new roman", 15),
                                       state="readonly",
                                       justify=CENTER)
        self.cmb_Estado['values'] = ('Select',
                                     'AC',
                                     'AL',
                                     'AM',
                                     'AP',
                                     'BA',
                                     'CE',
                                     'DF',
                                     'ES',
                                     'GO',
                                     'MA',
                                     'MG',
                                     'MS',
                                     'MT',
                                     'PA',
                                     'PB',
                                     'PE',
                                     'PI',
                                     'PR',
                                     'RJ',
                                     'RN',
                                     'RO',
                                     'RR',
                                     'RS',
                                     'SC',
                                     'SE',
                                     'SP',
                                     'TO')
        self.cmb_Estado.current(0)
        self.cmb_Estado.place(x=331, y=420, width=150)

        # - Cidade
        ## [TEXTO] Cidade
        textCidade = Label(self.root,
                           text="Cidade*",
                           font=("times new roman", 14, "bold"),
                           bg="#EEF1F2",
                           anchor="w",
                           fg="black")
        textCidade.place(x=331, y=465, width=70, height=17)

        ## [INPUT] Cidade
        self.txt_cidade = Entry(self.root,
                                font=("times new roman", 15),
                                bg="lightgray")
        self.txt_cidade.place(x=331, y=490, width=300)

        # - Itens consultados
        ## [TEXTO] Item 1
        textItem1 = Label(self.root,
                          text="Item 1*",
                          font=("times new roman", 14, "bold"),
                          bg="#EEF1F2",
                          anchor="w",
                          fg="black")
        textItem1.place(x=788, y=255, width=70, height=17)

        ## [INPUT] Item 1
        itemVar1 = StringVar()
        self.txt_Item1 = Entry(self.root,
                               font=("times new roman", 15),
                               bg="lightgray",
                               textvariable=itemVar1)
        itemVar1.trace("w", lambda name, index, mode, sv=itemVar1: entryUpdateItem(self.txt_Item1))
        self.txt_Item1.place(x=788, y=280, width=100)

        ## [TEXTO] Item 2
        textItem2 = Label(self.root,
                          text="Item 2",
                          font=("times new roman", 14, "bold"),
                          bg="#EEF1F2",
                          anchor="w",
                          fg="black")
        textItem2.place(x=933, y=255, width=60, height=17)

        ## [INPUT] Item 2
        itemVar2 = StringVar()
        self.txt_Item2 = Entry(self.root,
                               font=("times new roman", 15),
                               bg="lightgray",
                               textvariable=itemVar2)
        itemVar2.trace("w", lambda name, index, mode, sv=itemVar2: entryUpdateItem(self.txt_Item2))
        self.txt_Item2.place(x=933, y=280, width=100)

        ## [TEXTO] Item 3
        textItem3 = Label(self.root,
                          text="Item 3",
                          font=("times new roman", 14, "bold"),
                          bg="#EEF1F2",
                          anchor="w",
                          fg="black")
        textItem3.place(x=1078, y=255, width=60, height=17)

        ## [INPUT] Item 3
        itemVar3 = StringVar()
        self.txt_Item3 = Entry(self.root,
                               font=("times new roman", 15),
                               bg="lightgray",
                               textvariable=itemVar3)
        itemVar3.trace("w", lambda name, index, mode, sv=itemVar3: entryUpdateItem(self.txt_Item3))
        self.txt_Item3.place(x=1078, y=280, width=100)

        ## [TEXTO] Item 4
        textItem4 = Label(self.root,
                          text="Item 4",
                          font=("times new roman", 14, "bold"),
                          bg="#EEF1F2",
                          anchor="w",
                          fg="black")
        textItem4.place(x=788, y=325, width=60, height=17)

        ## [INPUT] Item 4
        itemVar4 = StringVar()
        self.txt_Item4 = Entry(self.root,
                               font=("times new roman", 15),
                               bg="lightgray",
                               textvariable=itemVar4)
        itemVar4.trace("w", lambda name, index, mode, sv=itemVar4: entryUpdateItem(self.txt_Item4))
        self.txt_Item4.place(x=788, y=350, width=100)

        ## [TEXTO] Item 5
        textItem5 = Label(self.root,
                          text="Item 5",
                          font=("times new roman", 14, "bold"),
                          bg="#EEF1F2",
                          anchor="w",
                          fg="black")
        textItem5.place(x=933, y=325, width=60, height=17)

        ## [INPUT] Item 5
        itemVar5 = StringVar()
        self.txt_Item5 = Entry(self.root,
                               font=("times new roman", 15),
                               bg="lightgray",
                               textvariable=itemVar5)
        itemVar5.trace("w", lambda name, index, mode, sv=itemVar5: entryUpdateItem(self.txt_Item5))
        self.txt_Item5.place(x=933, y=350, width=100)

        ## [TEXTO] Item 6
        textItem6 = Label(self.root,
                          text="Item 6",
                          font=("times new roman", 14, "bold"),
                          bg="#EEF1F2",
                          anchor="w",
                          fg="black")
        textItem6.place(x=1078, y=325, width=60, height=17)

        ## [INPUT] Item 6
        itemVar6 = StringVar()
        self.txt_Item6 = Entry(self.root,
                               font=("times new roman", 15),
                               bg="lightgray",
                               textvariable=itemVar6)
        itemVar6.trace("w", lambda name, index, mode, sv=itemVar6: entryUpdateItem(self.txt_Item6))
        self.txt_Item6.place(x=1078, y=350, width=100)

        # - Observações
       ## [TEXTO] Observações
        textObs = Label(self.root,
                        text="Observações",
                        font=("times new roman", 14, "bold"),
                        bg="#EEF1F2",
                        anchor="w",
                        fg="black")
        textObs.place(x=788, y=395, width=125, height=17)

        ## [INPUT] Observações
        self.txt_Obs = tkinter.Text(self.root,
                                    font=("times new roman", 15),
                                    bg="lightgray")
        self.txt_Obs.place(x=788, y=415, width=390, height=105)

        # Register Button ----------------------------------------------------
        img_register = Image.open("D:/Projetos/formulario_consultas/img/btn.png").resize((250, 60))
        self.btn_img = ImageTk.PhotoImage(img_register)
        btn_register = Button(self.root,
                              image=self.btn_img,
                              bd=0,
                              cursor="hand2",
                              bg="#EEF1F2",
                              command=self.register)
        btn_register.place(x=930, y=575, height=60, width=250)

        # Texto Informativo
        info_head = Label(self.root,
                          text="Utilização",
                          font=("times new roman", 15, "bold"),
                          bg="#EEF1F2",
                          anchor="w")
        info_head.place(x=35, y=142, height=30, width=210)

        self.info_txt = Message(self.root,
                                text="Os dados inseridos aqui, deverão \n"
                                     "ser a respeito de consultas de \n"
                                     "itens feito por clientes, para \n"
                                     "que possa ser identificado \n"
                                     "possíveis falhas de atendimento \n"
                                     "em regiões específicas.\n"
                                     "O relatório de consultas pode ser \n"
                                     "acessado no PBI através do  \n"
                                     "botão abaixo.",
                                font=("times new roman", 12),
                                bg="#EEF1F2",
                                anchor='nw')
        self.info_txt.place(x=33, y=175, height=300, width=210)

        # PBI Button
        img_pbi = Image.open("D:/Projetos/formulario_consultas/img/pbi.jpg").resize((220, 65))
        self.btn_img_pbi = ImageTk.PhotoImage(img_pbi)
        btn_pbi = Button(self.root,
                         image=self.btn_img_pbi,
                         bd=0,
                         cursor="hand2",
                         bg="#F2C80F",
                         command=self.openPbi)
        btn_pbi.place(x=30, y=600, height=65, width=220)

        def entryUpdateItem(entry):
            text = entry.get()
            if len(text) in (3, 3):
                entry.insert(END, '.')
                entry.icursor(len(text) + 1)
            elif len(text) not in (3, 6):
                if not text[-1].isdigit():
                    entry.delete(0, END)
                    entry.insert(0, text[:-1])
            if len(text) > 7:
                entry.delete(0, END)
                entry.insert(0, text[:7])

    def openPbi(self):
        webbrowser.open(url, new=1)

    def checkEmail(self, val):
        if re.search(regex, val):
            self.wdgLst.configure(text='')
            return True
        else:
            self.wdgLst.configure(text='E-mail Inválido')
            return False

    def register(self):
        if self.txt_Name.get() == '' or self.txt_Email.get() == '' or self.cmb_Estado.get() == 'Select' or self.wdgLst.cget(
                'text') != '' or self.txt_cidade.get() == '' or self.txt_Item1.get() == '':
            messagebox.showerror("!!!! Erro !!!!",
                                 "Há campos obrigatórios não preenchidos ou preenchidos de forma incorreta! Por favor verifique e tente novamente.",
                                 parent=self.root)
        else:
            server = 'DESKTOP-BNQAB8B\SQLEXPRESS' 
            database = 'ContosoRetailDW' 
            username = 'teste' 
            password = 'teste123456' 
            cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';Trusted_Connection=yes;')
            cursor = cnxn.cursor()
                            
            values = ([self.infoDataHoje,
            self.infoUser,
            self.txt_Name.get(),
            self.txt_Email.get(),
            self.txt_Telefone.get(),
            self.cmb_Estado.get(),
            self.txt_cidade.get(),
            self.txt_Item1.get(),
            self.txt_Item2.get(),
            self.txt_Item3.get(),
            self.txt_Item4.get(),
            self.txt_Item5.get(),
            self.txt_Item6.get(),
            self.txt_Obs.get(1.0, END)])

            cursor.execute('''INSERT INTO DimRegistrosConsultas (data, usuario, nome, email, telefone, estado, cidade, item1, item2, item3, item4, item5, item6, obs)
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);''', 
                             values)

            cnxn.commit()

            messagebox.showinfo("Registrado!", "O registro de consulta foi realizado!", parent=self.root)
            self.txt_Name.delete(0, END)
            self.txt_Email.delete(0, END)
            self.txt_Telefone.delete(0, END)
            self.txt_cidade.delete(0, END)
            self.txt_Item1.delete(0, END)
            self.txt_Item2.delete(0, END)
            self.txt_Item3.delete(0, END)
            self.txt_Item4.delete(0, END)
            self.txt_Item5.delete(0, END)
            self.txt_Item6.delete(0, END)
            self.txt_Obs.delete(1.0, END)
            self.cmb_Estado.current(0)


root = Tk()
regex = '^[a-z0-9]+[\._]?[a-z0-9]+[@]\w+[.]\w{2,8}$'
url = ''
obj = Register(root)
root.mainloop()
