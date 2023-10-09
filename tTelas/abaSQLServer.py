import pyodbc
from tkinter import messagebox
import ttkbootstrap as ttk
import customtkinter as ctk

from tTelas.importadorExcel import frame4

#Servidor
labelServidorSQLServer = ttk.Label(frame4, text="IP Servidor", font=ctk.CTkFont(size=10, weight="bold"))
labelServidorSQLServer.place(relx=0.1, rely=0.23, relwidth=0.4, relheight=0.2)
tituloServidorSQLServer = ttk.Entry(frame4)
tituloServidorSQLServer.place(relx=0.45, rely=0.25, relwidth=0.45, relheight=0.13)

labelDatabaseSQLServer = ttk.Label(frame4, text="Database", font=ctk.CTkFont(size=10, weight="bold"))
labelDatabaseSQLServer.place(relx=0.1, rely=0.43, relwidth=0.4, relheight=0.2)
tituloDatabaseSQLServer = ttk.Entry(frame4)
tituloDatabaseSQLServer.place(relx=0.45, rely=0.45, relwidth=0.45, relheight=0.13)

def testarConexaoSQLServer():
    if tituloServidorSQLServer.get() == '' or tituloDatabaseSQLServer.get() == '':
        messagebox.showwarning("Atenção", "Preencha o servidor e a base para poder testar a conexão.")
    else:
        try:
            conn = pyodbc.connect("DRIVER={SQL Server};SERVER=" + tituloServidorSQLServer.get() + ";DATABASE=" + tituloDatabaseSQLServer.get() + ";Trusted_Connection=yes")
            conn.close()
            messagebox.showinfo("Sucesso", "Conexão realizada com sucesso!")
        except:
            messagebox.showerror("Falha", "Falha ao tentar se conectar com o banco.")

testarConexao = ttk.Button(frame4, text='Testar Conexão', command=testarConexaoSQLServer)
testarConexao.place(relx=0.3, rely=0.83, relwidth=0.45, relheight=0.1)