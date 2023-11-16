import ttkbootstrap as ttk
import customtkinter as ctk

from tTelas.importadorExcel import frame5

label01 = ttk.Label(frame5, text="1 - Selecionar quais tipos de importação deseja realizar;", font=ctk.CTkFont(size=10, weight="bold"))
label01.place(relx=0.05, rely=0.05, relheight=0.05)
label011 = ttk.Label(frame5, text="1.1 - Se selecionado algum banco de dados diferente do \nSQLite, necessário informar os dados de conexão;", font=ctk.CTkFont(size=10, weight="bold"))
label011.place(relx=0.1, rely=0.1, relheight=0.1)
label02 = ttk.Label(frame5, text="2 - Para o MySQL, necessário informar todos \nos parâmetros de conexão;", font=ctk.CTkFont(size=10, weight="bold"))
label02.place(relx=0.05, rely=0.2, relheight=0.1)
label021 = ttk.Label(frame5, text="2.1 - As colunas serão criadas como VARCHAR(255);", font=ctk.CTkFont(size=10, weight="bold"))
label021.place(relx=0.1, rely=0.3, relheight=0.05)
label03 = ttk.Label(frame5, text="3 - Para o SQLServer, será utilizado o Trusted_Connection=yes, \nsendo necessário informar apenas o IP e a Base;", font=ctk.CTkFont(size=10, weight="bold"))
label03.place(relx=0.05, rely=0.35, relheight=0.1)
label031 = ttk.Label(frame5, text="3.1 - As colunas serão criadas como NVARCHAR(255);", font=ctk.CTkFont(size=10, weight="bold"))
label031.place(relx=0.1, rely=0.45, relheight=0.05)
label04 = ttk.Label(frame5, text="4 - Selecionar o arquivo a ser importado;", font=ctk.CTkFont(size=10, weight="bold"))
label04.place(relx=0.05, rely=0.5, relheight=0.05)
label041 = ttk.Label(frame5, text="4.1 - Nesta versão, apenas será considerado uma aba do arquivo;", font=ctk.CTkFont(size=10, weight="bold"))
label041.place(relx=0.1, rely=0.55, relheight=0.05)
label042 = ttk.Label(frame5, text="4.2 - O nome definido na aba do arquivo será o nome da tabela;", font=ctk.CTkFont(size=10, weight="bold"))
label042.place(relx=0.1, rely=0.6, relheight=0.05)
label05 = ttk.Label(frame5, text="5 - Durante o processamento dos dados, serão executadas algumas \nvalidações para previnir alguns possíveis erros;", font=ctk.CTkFont(size=10, weight="bold"))
label05.place(relx=0.05, rely=0.65, relheight=0.1)
label06 = ttk.Label(frame5, text="6 - O arquivo a ser importado deve seguir o padrão \ndo arquivo contido na pasta cConsumir;", font=ctk.CTkFont(size=10, weight="bold"))
label06.place(relx=0.05, rely=0.75, relheight=0.1)
label07 = ttk.Label(frame5, text="7 - Sempre será gerado o arquivo CSV;", font=ctk.CTkFont(size=10, weight="bold"))
label07.place(relx=0.05, rely=0.85, relheight=0.1)