from sqlalchemy.engine import URL, create_engine
from datetime import date, datetime, timedelta
from tkinter import filedialog, ttk, messagebox
from tkinter import *
import tkinter as tk
import requests
import pandas as pd
import json
import pyodbc
import openpyxl
import os


# 2 -conexao
connection_string = "DRIVER={SQL Server Native Client 11.0};SERVER=w2019.hausz.com.br;DATABASE=HauszMapa;UID=Aplicacao;PWD=S3nh4Apl!caca0"
connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})

engine = create_engine(connection_url)
conn = engine.connect()
print("Conexão Bem Sucedida")

# 3 -init
horas = timedelta(hours=-24)
hoje = datetime.today()
resultado = hoje + horas
horas2 = timedelta(hours=-48)
resultado2 = hoje + horas2
horas3 = timedelta(hours=72)
Reativacao1 = hoje + horas3
horas3 = timedelta(hours=120)
Reativacao2 = hoje + horas3
horas4 = timedelta(hours=360)
Reativacao3 = hoje + horas3



lista_disparo = ''


hojeRelatorios = date.today()
json = json
dict_messages = ['messages']
APIUrl = 'https://eu41.chat-api.com/instance397901/'
token = 'gal2y0wmxp7301lz'

Leads24h = f"""SELECT  C.Franquia,CO.NomeColaborador, U.Celular ,C.QntLeads, C.LeadsParados
FROM (
SELECT DISTINCT --C.NomeColaborador,CU.Celular,
UN.IdUnidade, UN.Nome Franquia,
COUNT(LF.Celular) QntLeads,-- OVER (PARTITION BY UN.IdUnidade) QntLeads,
COUNT(CASE WHEN LF.IdPosicao = 1 THEN 1 END) LeadsParados--OVER (PARTITION BY UN.IdUnidade) LeadsParados
--,CASE WHEN LF.bitLeadInteragido = 0 THEN 'Não interagiu' ELSE 'Interagiu' END LeadInteragido
--,FORMAT(LF.DataInserido, 'd', 'pt-br') DataInserido
FROM HauszMapa.Wpp.LeadFranquia LF
INNER JOIN HauszLogin.Cadastro.Unidade UN ON UN.IdUnidade = LF.IdUnidade
LEFT JOIN HauszLogin.Cadastro.Colaborador C ON C.IdColaborador = LF.IdColaborador
LEFT JOIN HauszLogin.Cadastro.Usuario CU ON CU.CpfCnpj = C.CpfCnpj
WHERE --bitLeadInteragido = 0
1=1
AND CAST(LF.DataInserido AS DATE) <= '{resultado}'
AND UN.IdUnidade IS NOT NULL AND C.IdColaborador IS NOT NULL
AND IdCampanha <> 164 
--and C.IdPerfilUsuario = 4
AND C.bitAtivo = 1
AND CU.bitAtivo = 1
AND UN.IdUnidade <> 1 
and UN.bitAtivo = 1
and Un.IdNivelLead = 0
and LF.bitAtivo = 1
GROUP BY UN.IdUnidade, UN.Nome
) AS C
INNER JOIN HauszLogin.Cadastro.Colaborador CO ON CO.IdUnidade = C.IdUnidade AND CO.IdPerfilUsuario = 4
inner join HauszLogin.Cadastro.Usuario U ON U.CpfCnpj = CO.CpfCnpj"""

Leads48h = f"""SELECT  C.Franquia,CO.NomeColaborador, U.Celular ,C.QntLeads, C.LeadsParados, c.IdNivelLead
FROM (
SELECT DISTINCT --C.NomeColaborador,CU.Celular,
UN.IdUnidade, UN.Nome Franquia, Un.IdNivelLead,
COUNT(LF.Celular) QntLeads,-- OVER (PARTITION BY UN.IdUnidade) QntLeads,
COUNT(CASE WHEN LF.IdPosicao = 1 THEN 1 END) LeadsParados--OVER (PARTITION BY UN.IdUnidade) LeadsParados
--,CASE WHEN LF.bitLeadInteragido = 0 THEN 'Não interagiu' ELSE 'Interagiu' END LeadInteragido
--,FORMAT(LF.DataInserido, 'd', 'pt-br') DataInserido
FROM Hauszmapa.Wpp.LeadFranquia LF
INNER JOIN HauszLogin.Cadastro.Unidade UN ON UN.IdUnidade = LF.IdUnidade
LEFT JOIN HauszLogin.Cadastro.Colaborador C ON C.IdColaborador = LF.IdColaborador
LEFT JOIN HauszLogin.Cadastro.Usuario CU ON CU.CpfCnpj = C.CpfCnpj
LEFT JOIN HauszLogin.Cadastro.UnidadeBit UB on UB.IdUnidade = Un.IdUnidade
WHERE --bitLeadInteragido = 0
1=1
AND CAST(LF.DataInserido AS DATE) <= '{resultado2}'
AND UN.IdUnidade IS NOT NULL AND C.IdColaborador IS NOT NULL
AND IdCampanha <> 164 
--and C.IdPerfilUsuario = 4
AND C.bitAtivo = 1
AND CU.bitAtivo = 1
AND UN.IdUnidade <> 1 
AND UB.bitCampanhaBot = 1
and UN.bitAtivo = 1
and LF.bitAtivo = 1
GROUP BY UN.IdUnidade, UN.Nome, Un.IdNivelLead
) AS C
INNER JOIN HauszLogin.Cadastro.Colaborador CO ON CO.IdUnidade = C.IdUnidade AND CO.IdPerfilUsuario = 4
inner join HauszLogin.Cadastro.Usuario U ON U.CpfCnpj = CO.CpfCnpj"""

Reativacao = f"""SELECT  C.Franquia,CO.NomeColaborador, U.Celular ,C.QntLeads, C.LeadsParados, c.IdNivelLead
FROM (
SELECT DISTINCT --C.NomeColaborador,CU.Celular,
UN.IdUnidade, UN.Nome Franquia, UN.IdNivelLead,
COUNT(LF.Celular) QntLeads,-- OVER (PARTITION BY UN.IdUnidade) QntLeads,
COUNT(CASE WHEN LF.IdPosicao = 1 THEN 1 END) LeadsParados--OVER (PARTITION BY UN.IdUnidade) LeadsParados
--,CASE WHEN LF.bitLeadInteragido = 0 THEN 'Não interagiu' ELSE 'Interagiu' END LeadInteragido
--,FORMAT(LF.DataInserido, 'd', 'pt-br') DataInserido
FROM HauszMapa.Wpp.LeadFranquia LF
INNER JOIN HauszLogin.Cadastro.Unidade UN ON UN.IdUnidade = LF.IdUnidade
LEFT JOIN HauszLogin.Cadastro.Colaborador C ON C.IdColaborador = LF.IdColaborador
LEFT JOIN HauszLogin.Cadastro.Usuario CU ON CU.CpfCnpj = C.CpfCnpj
LEFT JOIN HauszLogin.Cadastro.UnidadeBit UB on UB.IdUnidade = Un.IdUnidade
WHERE --bitLeadInteragido = 0
1=1
AND CAST(LF.DataInserido AS DATE) <= '{resultado}'
AND UN.IdUnidade IS NOT NULL AND C.IdColaborador IS NOT NULL
AND IdCampanha <> 164 
--and C.IdPerfilUsuario = 4
AND UN.IdNivelLead > 0
AND UB.bitCampanhaBot = 0
AND C.bitAtivo = 1
AND CU.bitAtivo = 1
AND UN.IdUnidade <> 1 
and UN.bitAtivo = 1
and LF.bitAtivo = 1
--AND un.Nome like '%Adamantina%'
GROUP BY UN.IdUnidade, UN.Nome, UN.IdNivelLead
) AS C
INNER JOIN HauszLogin.Cadastro.Colaborador CO ON CO.IdUnidade = C.IdUnidade AND CO.IdPerfilUsuario = 4
inner join HauszLogin.Cadastro.Usuario U ON U.CpfCnpj = CO.CpfCnpj"""

#tabelas

df = pd.read_sql_query(Leads24h, conn)
listaContato = df[['Franquia', 'NomeColaborador', 'Celular', 'LeadsParados']]
listaContato = listaContato.query('LeadsParados > 0')

df2 = pd.read_sql_query(Leads48h, conn)
listaContato2 = df2[['Franquia', 'NomeColaborador', 'Celular', 'LeadsParados']]
listaContato2 = listaContato2.query('LeadsParados > 0')

df3 = pd.read_sql_query(Reativacao, conn)
listaContato3 = df3[['Franquia', 'NomeColaborador', 'Celular', 'LeadsParados']]
listaContato3 = listaContato3.query('LeadsParados == 0')

def Salvar_planilha():
    dirpath = filedialog.askdirectory()
    listaContato.to_excel(f'{dirpath}/Lista_de_Aviso_24h.xlsx')
    listaContato2.to_excel(f'{dirpath}/Travamento_de_campanha_48h.xlsx')
    listaContato3.to_excel(f'{dirpath}/Lista_de_Reativação.xlsx')
    teste = messagebox.showinfo("Basic Example", "Download Concluido")

# BTN 4 Reativação de Campanha /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

def disparo_ReativarCampanha():
    arquivo = filedialog.askopenfilename()
    negativo = len(arquivo) - 4

    if len(arquivo) > 0 and arquivo[negativo:] == 'xlsx':

        Tela4 = Tk()  
        Tela4.title("Bot Dedo-Duro")

        larguta_tela = Tela4.winfo_screenwidth()
        altura_tela =  Tela4.winfo_screenmmheight()

        largura = 800
        altura = 400
        posix = larguta_tela/2 - largura/2
        posiy = altura_tela/2 - altura/6

        Tela4.geometry("%dx%d+%d+%d" % (largura,altura,posix,posiy))

        btnDispara = Button(Tela4, text="Disparar", command=send_message2)
        btnDispara.place(x= 40, y=315, height= 50, width= 200)

        btnVoltar = Button(Tela4, text="sair", command=Tela4.destroy)
        btnVoltar.place(x= 560, y=315, height= 50, width= 200)

        Tabela = ttk.Treeview(Tela4, selectmode="browse", columns=("Unidade", "Gestor", "Celular", "Leads_Parados", "IdNivelLead"), show='headings')
        Tabela.place(x=100, y=80, height=200)

        Tabela.column('Unidade', minwidth=0, width=150)
        Tabela.column('Gestor', minwidth=0, width=160)
        Tabela.column('Celular', minwidth=0, width=90)
        Tabela.column('Leads_Parados', minwidth=0, width=70)
        Tabela.column('IdNivelLead', minwidth=0, width=80)

        Tabela.heading('Unidade', text='Unidade')
        Tabela.heading('Gestor', text='Gestor')
        Tabela.heading('Celular', text='Celular')
        Tabela.heading('Leads_Parados', text='Leads_Parados')
        Tabela.heading('IdNivelLead', text='IdNivelLead')

        tabela_reativacao = pd.read_excel(f'{arquivo}')

        lista_disparo = tabela_reativacao
        lista_disparo.to_excel('lista_disparo.xlsx')
        print('IdNivelLead' in tabela_reativacao.columns)
        if 'IdNivelLead' in tabela_reativacao.columns == True:

            for index, row in tabela_reativacao.iterrows():
                Unidade = row["Franquia"]
                Gestor = row['NomeColaborador']
                Celular = row['Celular']
                LeadsParados = row['LeadsParados']
                Nivel = row['IdNivelLead']
                
                Tabela.insert("", "end", values=(Unidade, Gestor, Celular,LeadsParados, Nivel))
        else:
            Erro = messagebox.showerror("showerror", "A planilha adionada não corresponde ao padrão suportado. Verifique no 'Readme' o padrão esperado")
            Tela4.destroy()
            Tela.destroy()
            exec(open("engine.py").read())

    elif arquivo[negativo:] != 'xlsx' and len(arquivo) > 0:
        Erro = messagebox.showerror("showerror", "Arquivo Não Suportado")
        exec(open("engine.py").read())
    elif arquivo[negativo:] != 'xlsx' and len(arquivo) > 0:
        Erro = messagebox.showerror("showerror", "Nenhum Arquivo Encontrado")
        exec(open("engine.py").read())
    
def importa_excel_ReativarCampanha():
    Importa_arquivo = Tk()
    Importa_arquivo.title("Bot Dedo-Duro")

    larguta_tela = Importa_arquivo.winfo_screenwidth()
    altura_tela =  Importa_arquivo.winfo_screenmmheight()

    largura = 300
    altura = 150
    posix = larguta_tela/2 - largura/2
    posiy = altura_tela/2 - altura/8

    Importa_arquivo.geometry("%dx%d+%d+%d" % (largura,altura,posix,posiy))

    btnImportaArq = Button(Importa_arquivo, text="Importar Lista Disparo", command= lambda: [Importa_arquivo.destroy(), disparo_travamento_de_campanha()])
    btnImportaArq.place(x= 50, y=50, height= 50, width= 200)

def send_message4():
    lista_disparo = pd.read_excel('lista_disparo.xlsx')

    messagebox.showwarning("showwarning", "Disparo em andamento, por favor aguarde")

    total = len(lista_disparo.index)
    calc = 0
    for index, row in lista_disparo.iterrows():
            celular = row['Celular']
            Franquia = row['Franquia']
            Nivel = row['IdNivelLead']
            Leads = row['LeadsParados']

            calc =+ 1

            if Nivel == 1:
                engine.execute(f"""update hauszlogin.Cadastro.UnidadeBit set DataTerminoPunicaoLeads = '{Reativacao1}' where IdUnidade in (select IdUnidade from HauszLogin.Cadastro.Unidade where nome like '{Franquia}')'""")

                chatId = f'55{celular}@c.us'
                # chatId = '5519994790200@c.us'
                texto = f"""[MENSAGEM AUTOMÁTICA]

*Sua campanha foi desativada*

Identificamos que a sua franquia {Franquia} possui *{Leads} leads desatualizado* na coluna do CRM "Leads novos" a mais de *48h úteis*. 😬
Sua campanha está pausada temporariamente até que os leads sejam atualizados. 
O prazo para *reativar a campanha é de 3 dias úteis após a atualização do CRM*. 

É muito importante atualizar o status desses leads, manter sua carteirização em dia, não perder o time do cliente e para manter sua campanha ativa.

Conte conosco, Boas vendas 🚀

Dúvidas? Contate seu GC 😉"""

                data = {"chatId": chatId,
                        "body" : texto}
                answer = send_requests('sendMessage', data)

            elif Nivel == 2:
                engine.execute(f"""update hauszlogin.Cadastro.UnidadeBit set DataTerminoPunicaoLeads = '{Reativacao2}' where IdUnidade in (select IdUnidade from HauszLogin.Cadastro.Unidade where nome like '{Franquia}')'""")

                chatId = f'55{celular}@c.us'
                # chatId = '5519994790200@c.us'
                texto = f"""[MENSAGEM AUTOMÁTICA]

*Sua campanha foi desativada*

Identificamos que a sua franquia {Franquia} possui *{Leads} leads desatualizado* na coluna do CRM "Leads novos" a mais de *48h úteis*. 😬
Sua campanha está pausada temporariamente até que os leads sejam atualizados. 
O prazo para *reativar a campanha é de 5 dias úteis após a atualização do CRM*. 

É muito importante atualizar o status desses leads, manter sua carteirização em dia, não perder o time do cliente e para manter sua campanha ativa.

Conte conosco, Boas vendas 🚀

Dúvidas? Contate seu GC 😉"""

                data = {"chatId": chatId,
                        "body" : texto}
                answer = send_requests('sendMessage', data)

            else:
                engine.execute(f"""update hauszlogin.Cadastro.UnidadeBit set DataTerminoPunicaoLeads = '{Reativacao3}' where IdUnidade in (select IdUnidade from HauszLogin.Cadastro.Unidade where nome like '{Franquia}')'""")

                chatId = f'55{celular}@c.us'
                # chatId = '5519994790200@c.us'
                texto = f"""[MENSAGEM AUTOMÁTICA]

*Sua campanha foi desativada*

Identificamos que a sua franquia {Franquia} possui *{Leads} leads desatualizado* na coluna do CRM "Leads novos" a mais de *48h úteis*. 😬
Sua campanha está pausada temporariamente até que os leads sejam atualizados. 
O prazo para *reativar a campanha é de 15 dias úteis após a atualização do CRM*. 

É muito importante atualizar o status desses leads, manter sua carteirização em dia, não perder o time do cliente e para manter sua campanha ativa.

Conte conosco, Boas vendas 🚀

Dúvidas? Contate seu GC 😉"""

                data = {"chatId": chatId,
                        "body" : texto}
                answer = send_requests('sendMessage', data)

            if calc == total:
                messagebox.showinfo("showinfo", "Enviado com sucesso")
  
                Tela.destroy()
                exit()

# BTN 3 Travamento de Campanha 48h /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

def disparo_travamento_de_campanha():
    arquivo = filedialog.askopenfilename()
    negativo = len(arquivo) - 4

    if len(arquivo) > 0 and arquivo[negativo:] == 'xlsx':

        Tela4 = Tk()  
        Tela4.title("Bot Dedo-Duro")

        larguta_tela = Tela4.winfo_screenwidth()
        altura_tela =  Tela4.winfo_screenmmheight()

        largura = 800
        altura = 400
        posix = larguta_tela/2 - largura/2
        posiy = altura_tela/2 - altura/6

        Tela4.geometry("%dx%d+%d+%d" % (largura,altura,posix,posiy))

        btnDispara = Button(Tela4, text="Disparar", command=send_message2)
        btnDispara.place(x= 40, y=315, height= 50, width= 200)

        btnVoltar = Button(Tela4, text="sair", command=Tela4.destroy)
        btnVoltar.place(x= 560, y=315, height= 50, width= 200)

        Tabela = ttk.Treeview(Tela4, selectmode="browse", columns=("Unidade", "Gestor", "Celular", "Leads_Parados", "IdNivelLead"), show='headings')
        Tabela.place(x=100, y=80, height=200)

        Tabela.column('Unidade', minwidth=0, width=150)
        Tabela.column('Gestor', minwidth=0, width=160)
        Tabela.column('Celular', minwidth=0, width=90)
        Tabela.column('Leads_Parados', minwidth=0, width=70)
        Tabela.column('IdNivelLead', minwidth=0, width=80)

        Tabela.heading('Unidade', text='Unidade')
        Tabela.heading('Gestor', text='Gestor')
        Tabela.heading('Celular', text='Celular')
        Tabela.heading('Leads_Parados', text='Leads_Parados')
        Tabela.heading('IdNivelLead', text='IdNivelLead')

        Tabela_travamento = pd.read_excel(f'{arquivo}')

        lista_disparo = Tabela_travamento
        lista_disparo.to_excel('Icon_Hausz_72x72.ico')
        print('IdNivelLead' in Tabela_travamento.columns)
        if 'IdNivelLead' in Tabela_travamento.columns == True:

            for index, row in Tabela_travamento.iterrows():
                Unidade = row["Franquia"]
                Gestor = row['NomeColaborador']
                Celular = row['Celular']
                LeadsParados = row['LeadsParados']
                Nivel = row['IdNivelLead']
                
                Tabela.insert("", "end", values=(Unidade, Gestor, Celular,LeadsParados ))
        else:
            Erro = messagebox.showerror("showerror", "A planilha adionada não corresponde ao padrão suportado. Verifique no Readme.txt o padrão esperado")
            Tela4.destroy()
            Tela.destroy()
            exec(open("engine.py").read())

    elif arquivo[negativo:] != 'xlsx' and len(arquivo) > 0:
        Erro = messagebox.showerror("showerror", "Arquivo Não Suportado")
        exec(open("engine.py").read())
    elif arquivo[negativo:] != 'xlsx' and len(arquivo) > 0:
        Erro = messagebox.showerror("showerror", "Nenhum Arquivo Encontrado")
        exec(open("engine.py").read())
    
def importa_excel_travamento():
    Importa_arquivo = Tk()
    Importa_arquivo.title("Bot Dedo-Duro")
    larguta_tela = Importa_arquivo.winfo_screenwidth()
    altura_tela =  Importa_arquivo.winfo_screenmmheight()

    largura = 300
    altura = 150
    posix = larguta_tela/2 - largura/2
    posiy = altura_tela/2 - altura/8

    Importa_arquivo.geometry("%dx%d+%d+%d" % (largura,altura,posix,posiy))

    btnImportaArq = Button(Importa_arquivo, text="Importar Lista Disparo", command= lambda: [Importa_arquivo.destroy(), disparo_travamento_de_campanha()])
    btnImportaArq.place(x= 50, y=50, height= 50, width= 200)

def send_message3():
    lista_disparo = pd.read_excel('lista_disparo.xlsx')

    messagebox.showwarning("showwarning", "Disparo em andamento, por favor aguarde")

    total = len(lista_disparo.index)
    calc = 0
    for index, row in lista_disparo.iterrows():
            celular = row['Celular']
            Franquia = row['Franquia']
            Leads = row['LeadsParados']
            Nivel = row['IdNivelLead']
            calc =+ 1

            engine.execute(f"""update HauszLogin.Cadastro.UnidadeBit SET bitCampanhaBot = 0 where IdUnidade in (select IdUnidade from HauszLogin.Cadastro.Unidade where Nome like '{Franquia}')""")
            if Nivel <= 2:
                engine.execute(f"""update HauszLogin.Cadastro.Unidade SET IdNivelLead = IdNivelLead + 1 where Nome like '{Franquia}'""")
            else:
                print('Nivel Maximo')

            if Nivel == 1:

                chatId = f'55{celular}@c.us'
                # chatId = '5519994790200@c.us'
                texto = f"""[MENSAGEM AUTOMÁTICA]

*Ativação de campanha*

Identificamos que a sua franquia {Franquia} atualizou os leads do CRM. Iniciamos o processo de *ativação do seu Marketing*. Em até *3 dias úteis sua campanha estará ativa novamente*. 

Para não ficar sem leads não deixe de atualizar diariamente o seu CRM. 

Conte conosco, Boas vendas 🚀

Dúvidas? Contate seu GC 😉"""

                data = {"chatId": chatId,
                        "body" : texto}
                answer = send_requests('sendMessage', data)


            elif Nivel == 2:

                chatId = f'55{celular}@c.us'
                # chatId = '5519994790200@c.us'
                texto = f"""[MENSAGEM AUTOMÁTICA]

*Ativação de campanha*

Identificamos que a sua franquia {Franquia} atualizou os leads do CRM. Iniciamos o processo de *ativação do seu Marketing*. Em até *5 dias úteis sua campanha estará ativa novamente*. 

Para não ficar sem leads não deixe de atualizar diariamente o seu CRM. 

Conte conosco, Boas vendas 🚀

Dúvidas? Contate seu GC 😉"""

                data = {"chatId": chatId,
                        "body" : texto}
                answer = send_requests('sendMessage', data)

            else:
    
                chatId = f'55{celular}@c.us'
                # chatId = '5519994790200@c.us'
                texto = f"""[MENSAGEM AUTOMÁTICA]

*Ativação de campanha*

Identificamos que a sua franquia {Franquia} atualizou os leads do CRM. Iniciamos o processo de *ativação do seu Marketing*. Em até *15 dias úteis sua campanha estará ativa novamente*. 

Para não ficar sem leads não deixe de atualizar diariamente o seu CRM. 

Conte conosco, Boas vendas 🚀

Dúvidas? Contate seu GC 😉"""

                data = {"chatId": chatId,
                        "body" : texto}
                answer = send_requests('sendMessage', data)


            if calc == total:
                messagebox.showinfo("showinfo", "Enviado com sucesso")
                
                exit()

# BTN 2 Aviso 24h Não Inaugurados /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
def disparo_aviso_n_inaugurado():
    arquivo = filedialog.askopenfilename()
    negativo = len(arquivo) - 4

    if len(arquivo) > 0 and arquivo[negativo:] == 'xlsx':

        Tela3 = Tk()  
        Tela3.title("Bot Dedo-Duro")


        larguta_tela = Tela3.winfo_screenwidth()
        altura_tela =  Tela3.winfo_screenmmheight()

        largura = 800
        altura = 400
        posix = larguta_tela/2 - largura/2
        posiy = altura_tela/2 - altura/6

        Tela3.geometry("%dx%d+%d+%d" % (largura,altura,posix,posiy))

        btnDispara = Button(Tela3, text="Disparar", command=send_message2)
        btnDispara.place(x= 40, y=315, height= 50, width= 200)

        btnVoltar = Button(Tela3, text="sair", command=Tela3.destroy)
        btnVoltar.place(x= 560, y=315, height= 50, width= 200)

        Tabela = ttk.Treeview(Tela3, selectmode="browse", columns=("Unidade", "Gestor", "Celular", "Leads_Parados"), show='headings')
        Tabela.place(x=100, y=80, height=200)

        Tabela.column('Unidade', minwidth=0, width=150)
        Tabela.column('Gestor', minwidth=0, width=230)
        Tabela.column('Celular', minwidth=0, width=100)
        Tabela.column('Leads_Parados', minwidth=0, width=80)

        Tabela.heading('Unidade', text='Unidade')
        Tabela.heading('Gestor', text='Gestor')
        Tabela.heading('Celular', text='Celular')
        Tabela.heading('Leads_Parados', text='Leads_Parados')

        tabela24_n_inaugurada = pd.read_excel(f'{arquivo}')

        lista_disparo = tabela24_n_inaugurada
        lista_disparo.to_excel('Icon_Hausz_72x72.ico')

        for index, row in tabela24_n_inaugurada.iterrows():

            Unidade = row["Franquia"]
            Gestor = row['NomeColaborador']
            Celular = row['Celular']
            LeadsParados = row['LeadsParados']

            Tabela.insert("", "end", values=(Unidade, Gestor, Celular,LeadsParados ))
    
    elif arquivo[negativo:] != 'xlsx' and len(arquivo) > 0:
        Erro = messagebox.showerror("showerror", "Arquivo Não Suportado")
        exec(open("engine.py").read())
    else:
        Erro = messagebox.showerror("showerror", "Nenhum Arquivo Encontrado")
        exec(open("engine.py").read())
def importa_excel_Aviso():
    Importa_arquivo = Tk()
    Importa_arquivo.title("Bot Dedo-Duro")

    larguta_tela = Importa_arquivo.winfo_screenwidth()
    altura_tela =  Importa_arquivo.winfo_screenmmheight()

    largura = 300
    altura = 150
    posix = larguta_tela/2 - largura/2
    posiy = altura_tela/2 - altura/8

    Importa_arquivo.geometry("%dx%d+%d+%d" % (largura,altura,posix,posiy))

    btnImportaArq = Button(Importa_arquivo, text="Importar Lista Disparo", command= lambda: [disparo_aviso_n_inaugurado(), Importa_arquivo.destroy(), Tela.destroy()])
    btnImportaArq.place(x= 50, y=50, height= 50, width= 200)
def send_message2():
    lista_disparo = pd.read_excel('lista_disparo.xlsx')

    messagebox.showwarning("showwarning", "Disparo em andamento, por favor aguarde")

    total = len(lista_disparo.index)
    calc = 0
    for index, row in lista_disparo.iterrows():
            celular = row['Celular']
            Franquia = row['Franquia']
            Leads = row['LeadsParados']
            calc =+ 1

            chatId = f'55{celular}@c.us'
            # chatId = '5519994790200@c.us'
            texto = f"""[MENSAGEM AUTOMÁTICA]

Identificamos que a sua franquia {Franquia} possui *{Leads} leads desatualizados* na coluna do CRM "Leads novos". 😬
É muito importante atualizar o status desses leads, *manter sua carteirização em dia*, não perder o time do cliente e manter sua campanha ativa.

Conte conosco, Boas vendas 🚀

Dúvidas? Contate seu GC 😉"""

            data = {"chatId": chatId,
                    "body" : texto}
            answer = send_requests('sendMessage', data)

            if calc == total:
                messagebox.showinfo("showinfo", "Enviado com sucesso")

                exit()

# BTN 1 24h Inauguradas /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
def disparo_aviso():
    arquivo = filedialog.askopenfilename()
    negativo = len(arquivo) - 4

    if len(arquivo) > 0 and arquivo[negativo:] == 'xlsx':


        Tela2 = Tk()  
        Tela2.title("Bot Dedo-Duro")

        larguta_tela = Tela2.winfo_screenwidth()
        altura_tela =  Tela2.winfo_screenmmheight()

        largura = 800
        altura = 400
        posix = larguta_tela/2 - largura/2
        posiy = altura_tela/2 - altura/6

        Tela2.geometry("%dx%d+%d+%d" % (largura,altura,posix,posiy))

        btnDispara = Button(Tela2, text="Disparar", command=send_message1)
        btnDispara.place(x= 40, y=315, height= 50, width= 200)

        btnVoltar = Button(Tela2, text="sair", command=Tela2.destroy)
        btnVoltar.place(x= 560, y=315, height= 50, width= 200)

        Tabela = ttk.Treeview(Tela2, selectmode="browse", columns=("Unidade", "Gestor", "Celular", "Leads_Parados"), show='headings')
        Tabela.place(x=100, y=80, height=200)

        Tabela.column('Unidade', minwidth=0, width=150)
        Tabela.column('Gestor', minwidth=0, width=230)
        Tabela.column('Celular', minwidth=0, width=100)
        Tabela.column('Leads_Parados', minwidth=0, width=80)

        Tabela.heading('Unidade', text='Unidade')
        Tabela.heading('Gestor', text='Gestor')
        Tabela.heading('Celular', text='Celular')
        Tabela.heading('Leads_Parados', text='Leads_Parados')

        tabela24 = pd.read_excel(f'{arquivo}')
        print(tabela24)

        lista_disparo = tabela24
        lista_disparo.to_excel('lista_disparo.xlsx')

        for index, row in tabela24.iterrows():

            Unidade = row["Franquia"]
            Gestor = row['NomeColaborador']
            Celular = row['Celular']
            LeadsParados = row['LeadsParados']

            Tabela.insert("", "end", values=(Unidade, Gestor, Celular,LeadsParados ))

    
    elif arquivo[negativo:] != 'xlsx' and len(arquivo) > 0:
        Erro = messagebox.showerror("showerror", "Arquivo Não Suportado")
        exec(open("engine.py").read())

    else:
        Erro = messagebox.showerror("showerror", "Nenhum Arquivo Encontrado")
        exec(open("engine.py").read())

def importa_excel():
    Importa_arquivo = Tk()
    Importa_arquivo.title("Bot Dedo-Duro")
    larguta_tela = Importa_arquivo.winfo_screenwidth()
    altura_tela =  Importa_arquivo.winfo_screenmmheight()

    largura = 300
    altura = 150
    posix = larguta_tela/2 - largura/2
    posiy = altura_tela/2 - altura/8

    Importa_arquivo.geometry("%dx%d+%d+%d" % (largura,altura,posix,posiy))

    btnImportaArq = Button(Importa_arquivo, text="Importar Lista Disparo", command= lambda: [disparo_aviso(), Importa_arquivo.destroy(), Tela.destroy()])
    btnImportaArq.place(x= 50, y=50, height= 50, width= 200)

def send_message1():
    lista_disparo = pd.read_excel('lista_disparo.xlsx')

    teste = messagebox.showwarning("showwarning", "Disparo em andamento, por favor aguarde")

    total = len(lista_disparo.index)
    calc = 0
    for index, row in lista_disparo.iterrows():
            celular = row['Celular']
            Franquia = row['Franquia']
            Leads = row['LeadsParados']
            calc =+ 1

            chatId = f'55{celular}@c.us'
            # chatId = '5519994790200@c.us'
            texto = f"""[MENSAGEM AUTOMÁTICA]

*Dentro de 24h sua campanha será desativada*

Identificamos que a sua franquia {Franquia} possui *{Leads} leads desatualizados* na coluna do CRM "Leads novos". 😬
É muito importante atualizar o status desses leads, *manter sua carteirização em dia*, não perder o time do cliente e manter sua campanha ativa.

Conte conosco, Boas vendas 🚀

Dúvidas? Contate seu GC 😉"""

            data = {"chatId": chatId,
                    "body" : texto}
            answer = send_requests('sendMessage', data)

            if calc == total:
                messagebox.showinfo("showinfo", "Enviado com sucesso")
                exit()
                                
def send_requests( method, data):
    url = f"{APIUrl}{method}?token={token}"
    headers = {'Content-type': 'application/json'}
    answer = requests.post(url, data=json.dumps(data), headers=headers)
    return answer.json()

Tela = Tk()
Tela.title("Bot Dedo-Duro")

larguta_tela = Tela.winfo_screenwidth()
altura_tela =  Tela.winfo_screenmmheight()

largura = 600
altura = 480
posix = larguta_tela/2 - largura/2
posiy = altura_tela/2 - altura/6

Tela.geometry("%dx%d+%d+%d" % (largura,altura,posix,posiy))

label = Label(Tela, text="Bot Dedo-Duro", font='Poppins')
label.place(x= 200, y=50, height= 50, width= 200)


 # btn
btnDispara = Button(Tela, text="Disparo de Aviso (24h)", command=importa_excel)
btnDispara.place(x= 200, y=100, height= 50, width= 200)

btnDispara2 = Button(Tela, text="Disparar Aviso (não inaugurada)", command=importa_excel_Aviso)
btnDispara2.place(x= 200, y=180, height= 50, width= 200)

btnDispara3 = Button(Tela, text="Travamento de Campanha (48h)", command=importa_excel_travamento)
btnDispara3.place(x= 200, y=260, height= 50, width= 200)

btnDispara4 = Button(Tela, text="Reativação de Campanha", command=importa_excel_ReativarCampanha)
btnDispara4.place(x= 200, y=340, height= 50, width= 200)

btnDispara5 = Button(Tela, text="Baixar Relatorios", command=Salvar_planilha)
btnDispara5.place(x= 200, y=420, height= 50, width= 200)

Tela.mainloop()