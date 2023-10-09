import mysql.connector
from tkinter import messagebox
import ttkbootstrap as ttk
import customtkinter as ctk

from tTelas.importadorExcel import frame3

#Servidor
labelServidorMySQL = ttk.Label(frame3, text="IP Servidor", font=ctk.CTkFont(size=10, weight="bold"))
labelServidorMySQL.place(relx=0.1, rely=0.03, relwidth=0.4, relheight=0.2)
tituloServidorMySQL = ttk.Entry(frame3)
tituloServidorMySQL.place(relx=0.45, rely=0.05, relwidth=0.45, relheight=0.13)

labelUsuarioMySQL = ttk.Label(frame3, text="Usuário", font=ctk.CTkFont(size=10, weight="bold"))
labelUsuarioMySQL.place(relx=0.1, rely=0.23, relwidth=0.4, relheight=0.2)
tituloUsuarioMySQL = ttk.Entry(frame3)
tituloUsuarioMySQL.place(relx=0.45, rely=0.25, relwidth=0.45, relheight=0.13)

labelSenhaMySQL = ttk.Label(frame3, text="Senha", font=ctk.CTkFont(size=10, weight="bold"))
labelSenhaMySQL.place(relx=0.1, rely=0.43, relwidth=0.4, relheight=0.2)
tituloSenhaMySQL = ttk.Entry(frame3)
tituloSenhaMySQL.place(relx=0.45, rely=0.45, relwidth=0.45, relheight=0.13)

labelDatabaseMySQL = ttk.Label(frame3, text="Database", font=ctk.CTkFont(size=10, weight="bold"))
labelDatabaseMySQL.place(relx=0.1, rely=0.63, relwidth=0.4, relheight=0.2)
tituloDatabaseMySQL = ttk.Entry(frame3)
tituloDatabaseMySQL.place(relx=0.45, rely=0.65, relwidth=0.45, relheight=0.13)

def testarConexaoMySQL():
    if tituloServidorMySQL.get() == '' or tituloUsuarioMySQL.get() == '' or tituloSenhaMySQL.get() == '' or tituloDatabaseMySQL.get() == '':
        messagebox.showwarning("Atenção", "Preencha todos os campos para testar a conexão.")
    else:
        try:
            conn = mysql.connector.connect(
                host=tituloServidorMySQL.get(),
                user=tituloUsuarioMySQL.get(),
                password=tituloSenhaMySQL.get(),
                database=tituloDatabaseMySQL.get()
            )
            conn.close()
            messagebox.showinfo("Sucesso", "Conexão realizada com sucesso!")
        except:
            messagebox.showerror("Falha", "Falha ao tentar se conectar com o banco.")

testarConexao = ttk.Button(frame3, text='Testar Conexão', command=testarConexaoMySQL)
testarConexao.place(relx=0.3, rely=0.83, relwidth=0.45, relheight=0.1)