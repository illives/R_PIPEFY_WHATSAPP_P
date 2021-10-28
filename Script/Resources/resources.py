from cryptography.fernet import Fernet
from .log_resources import LogMixim
from datetime import date
import pyautogui as rb
import pywhatkit as py
import requests, json
import win32com.client
import pandas as pd
import sqlite3 
import os


class Credent:

    def __init__(self):
        self._pipelogin = None
        self._homedir = os.getcwd()[:-6]
        self.id_cards = []
        self.df_main = None
        self.lista_novas = []
        self.lista_aprovado = []
        self.lista_reprovado = []

    @property
    def mailpass(self):
        return self._mailpass

    @property
    def homedir(self):
        return self._homedir
    
    @property
    def maillogin(self):
        return self._maillogin
    
    @property
    def pipelogin(self):
        return self._pipelogin
    
    
    def credencias (self):
        """
        Extrair as Credencias do Arquivo de Credenciais
        """
        try:
            credent = f'{self._homedir}My_credent\\crednt.xlsx'
            VK = f'{self._homedir}My_credent\\VK.txt'
            SK = f'{self._homedir}My_credent\\SK.txt'
            key = Fernet.generate_key()
            file = open (VK, 'wb')
            file.write(key)
            file.close()
            file = open(SK, 'r')
            setenc = file.read()
            file.close()
            encodset = setenc.encode()
            f = Fernet(key)
            encrypted = f.encrypt(encodset)
            f2 = Fernet(key)
            decrypted = f2.decrypt(encrypted)
            token = decrypted.decode()
            xlApp = win32com.client.Dispatch("Excel.Application")
            xlwb = xlApp.workbooks.Open(credent, False, True, None,token)
            ws = xlwb.Worksheets(1)
            self._pipelogin = str(ws.Range("B2").value)
            self._maillogin = str(ws.Range("B3").value)
            self._mailpass = str(ws.Range("B4").value)
            xlwb.Close(True)
        except Exception as b:
            print(b)


class RequestModel(Credent, LogMixim):

    def listar_cards(self):
        """
        Listar todos os Card do PipeId.
        """
        try:
            url = "https://api.pipefy.com/graphql"
            payload = f"""{{\"query\":\"{{ allCards (pipeId:301462443) {{ edges {{ node {{ id current_phase {{name id}} fields {{ name  value }} comments {{  text }} }}}} pageInfo {{ endCursor }}}}}}\"}}"""
            headers = {
                    'authorization': f"Bearer {self._pipelogin}",
                    'content-type': "application/json"
                    }
            response = requests.request("POST", url, data=payload, headers=headers)
            dic=json.loads(response.text)
            endCursor = dic['data']['allCards']['pageInfo']['endCursor']
            for i in range(0,49):
                id_card = dic['data']['allCards']['edges'][i]['node']['id']
                current_phase = dic['data']['allCards']['edges'][i]['node']['current_phase']['name']
                fields = dic['data']['allCards']['edges'][i]['node']['fields']
                obs_analista = str(dic['data']['allCards']['edges'][i]['node']['comments'])
                nome_solicitante = ''
                placa = ''
                dta_agendamento = ''
                nome_fornecedor = ''
                end_fornecedor = ''
                telefone = ''
                for d in fields:
                    if d['name'] == 'Nome':
                        nome_solicitante = d['value']
                    elif d['name'] == 'PLACA':
                        placa = d['value']
                    elif d['name'] == 'SUGESTÃO DA MELHOR DATA':
                        dta_agendamento = d['value']
                    elif d['name'] == 'Nome do Fornecedor':
                        nome_fornecedor = d['value']
                    elif d['name'] == 'Endereço do Fornecedor':
                        end_fornecedor = d['value']
                    elif d['name'] == 'CELULAR':
                        telefone = d['value']
                    elif d['name'] == 'Observações do Analista' or d['name'] == 'Justificativa de Reprovação':
                        obs_analista1 = str(d['value'])
                        obs_analista = f'{obs_analista}_{obs_analista1}'
                dict = {'id_card': id_card, 'current_phase': current_phase, 'nome_solicitante': nome_solicitante, 'placa': placa, 'dta_agendamento': dta_agendamento, 'nome_fornecedor': nome_fornecedor, 'end_fornecedor': end_fornecedor, 'obs_analista': obs_analista, 'telefone': telefone}
                self.id_cards.append(dict)
            while True:
                try:
                    payload = f"""{{\"query\":\"{{ allCards (pipeId:301462443, after:\\"{endCursor}\\") {{ edges {{ node {{ id current_phase {{name id}} fields {{ name  value }} comments {{  text }} }}}} pageInfo {{ endCursor }}}}}}\"}}"""
                    response = requests.request("POST", url, data=payload, headers=headers)
                    dic=json.loads(response.text)
                    endCursor = dic['data']['allCards']['pageInfo']['endCursor']
                    for i in range(0,49):
                        id_card = dic['data']['allCards']['edges'][i]['node']['id']
                        current_phase = str(dic['data']['allCards']['edges'][i]['node']['current_phase']['name'])
                        fields = dic['data']['allCards']['edges'][i]['node']['fields']
                        obs_analista = str(dic['data']['allCards']['edges'][i]['node']['comments'])
                        nome_solicitante = ''
                        placa = ''
                        dta_agendamento = ''
                        nome_fornecedor = ''
                        end_fornecedor = ''
                        telefone = ''
                        for d in fields:
                            if d['name'] == 'Nome':
                                nome_solicitante = d['value']
                            elif d['name'] == 'PLACA':
                                placa = d['value']
                            elif d['name'] == 'SUGESTÃO DA MELHOR DATA':
                                dta_agendamento = d['value']
                            elif d['name'] == 'Nome do Fornecedor':
                                nome_fornecedor = d['value']
                            elif d['name'] == 'Endereço do Fornecedor':
                                end_fornecedor = d['value']
                            elif d['name'] == 'CELULAR':
                                telefone = d['value']
                            elif d['name'] == 'Observações do Analista' or d['name'] == 'Justificativa de Reprovação':
                                obs_analista1 = str(d['value'])
                                obs_analista = f'{obs_analista}_{obs_analista1}'
                            dict = {'id_card': id_card, 'current_phase': current_phase, 'nome_solicitante': nome_solicitante, 'placa': placa, 'dta_agendamento': dta_agendamento, 'nome_fornecedor': nome_fornecedor, 'end_fornecedor': end_fornecedor, 'obs_analista': obs_analista, 'telefone': telefone}
                            self.id_cards.append(dict)
                except Exception as b:
                    self.log_info(f'listar_cards_{len(self.id_cards)} Cards Consultados.')
                    break
        except Exception as b:
            self.log_erro(f'Listar_cards_ {b}')

class DataBase(RequestModel):

    def create_table_id_pipes(self):
        """
        Create table ID_PIPES_DB
        """
        conn = sqlite3.Connection(f'{self._homedir}My_database\\Dados.db')
        cursor = conn.cursor()
        cursor.execute('CREATE TABLE IF NOT EXISTS ID_PIPES_DB (ID_CARDS VARCHAR(50) PRIMARY KEY, STATUS_CURRENT_PHASE VARCHAR(50), NOME_SOLICITANTE VARCHAR(100), TEXTO_PLACA VARCHAR (20), DATA_AGENDAMENTO VARCHAR(50), NOME_FORNECEDOR VARCHAR (100), ENDERECO_FORNECEDOR VARCHAR(150), OBSERVACAO_ANALISTA VARCHAR (100), TELEFONE VARCHAR (50), STATUS_ENVIO_NOVA VARCHAR (20), STATUS_ENVIO_APROV_REJEI VARCHAR (20), DATA_ENVIO VARCHAR(50));')
        conn.commit()
        conn.close()
    
    def insert_new_cards(self):
        """
        Atualiza os Banco de Dados com os novos Cards Inseridos no Pipefy.\n
        Por definição todos os Cards estão com o status NAO_ENVIADO
        """
        try:
            self.create_table_id_pipes()
            conn = sqlite3.Connection(f'{self._homedir}My_database\\Dados.db')
            cursor = conn.cursor()
            valores = [(x['id_card'], x['current_phase'], x['nome_solicitante'], x['placa'], x['dta_agendamento'], x['nome_fornecedor'], x['end_fornecedor'], x['obs_analista'], x['telefone'], 'NAO_ENVIADO', 'NAO_ENVIADO','') for x in self.id_cards]
            cursor.executemany('INSERT OR IGNORE INTO ID_PIPES_DB VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);', valores)
            conn.commit()
            conn.close()
            self.log_info(f'DataBase_ Inseridas {cursor.rowcount} novos registros.')
        except Exception as b:
            self.log_erro(f'DataBase_ {b}')
    
    def update_atributos(self):
        """
        Atualizar campos de telefone, placa entre outros.
        """
        try:
            conn = sqlite3.Connection(f'{self._homedir}My_database\\Dados.db')
            cursor = conn.cursor()
            valores = [(x['current_phase'], x['nome_solicitante'], x['placa'], x['dta_agendamento'], x['nome_fornecedor'], x['end_fornecedor'], x['obs_analista'], x['telefone'], x['id_card']) for x in self.id_cards]
            cursor.executemany("""UPDATE OR IGNORE ID_PIPES_DB SET 
                                STATUS_CURRENT_PHASE = ?,
                                NOME_SOLICITANTE = ?,
                                TEXTO_PLACA = ?,
                                DATA_AGENDAMENTO = ?, 
                                NOME_FORNECEDOR = ?,
                                ENDERECO_FORNECEDOR = ?,
                                OBSERVACAO_ANALISTA = ?,
                                TELEFONE = ?
                                WHERE ID_CARDS = ?""", valores)
            conn.commit()
            conn.close()
            self.log_info(f'DataBase_ Atualizados {cursor.rowcount} registros')
        except Exception as b:
            self.log_erro(f'DataBase_ {b}')
    
    def update_status_envio_nova(self, idcard):
        """
        Atualizar status de novas requisicoes com data de envio.
        """
        try:
            tempo = str(date.today().strftime("%d-%m-%Y"))
            conn = sqlite3.Connection(f'{self._homedir}My_database\\Dados.db')
            cursor = conn.cursor()
            cursor.execute (f"""UPDATE ID_PIPES_DB SET
                            STATUS_ENVIO_NOVA = "ENVIADO",
                            DATA_ENVIO = "{tempo}"
                            WHERE ID_CARDS = "{idcard}"
                            """)
            conn.commit()
            conn.close()
            self.log_info(f'DataBase_ Atualizados {cursor.rowcount} Registros.')
        except Exception as b:
            self.log_erro(f'DataBase_ {b}')

    def update_status_envio_aprov_rejei(self, idcard):
        """
        Atualizar status aprovados ou rejeitados com data de envio.
        """
        try:
            tempo = str(date.today().strftime("%d-%m-%Y"))
            conn = sqlite3.Connection(f'{self._homedir}My_database\\Dados.db')
            cursor = conn.cursor()
            cursor.execute (f"""UPDATE ID_PIPES_DB SET
                            STATUS_ENVIO_APROV_REJEI = "ENVIADO",
                            DATA_ENVIO = "{tempo}"
                            WHERE ID_CARDS = "{idcard}"
                            """)
            conn.commit()
            conn.close()
            self.log_info(f'DataBase_ Atualizados {cursor.rowcount} Registros.')
        except Exception as b:
            self.log_erro(f'DataBase_ {b}')

class Report(DataBase):

    def relatorio_geral(self):
        """
        Data Frame com os STATUS_CURRENT_PHASE Agendado, Reprovado ou Novas Solicitacoes.
        """
        try:
            conn = sqlite3.Connection(f'{self._homedir}My_database\\Dados.db')
            self.df_main = pd.read_sql('SELECT * FROM ID_PIPES_DB WHERE STATUS_CURRENT_PHASE = "Reprovado ❌" OR STATUS_CURRENT_PHASE = "Agendado ✅" OR STATUS_CURRENT_PHASE = "NOVAS SOLICITAÇÕES"', conn)
            self.df_main['DATA_AGENDAMENTO'] = self.df_main['DATA_AGENDAMENTO'].astype('datetime64[ns]')
            self.df_main['DATA_AGENDAMENTO'] = self.df_main['DATA_AGENDAMENTO'].dt.strftime('%d/%m/%Y')
            self.log_info(f'Report_ Main Data Frame Criado.')
        except Exception as b:
            self.log_erro(f'Report_ {b}')
    
    def relatorio_novas_solicitacoes(self):
        """
        Criar Lista de Tuplas com (ID_CARD, TELEFONE, MENSAGEM) dos casos de novas solicitacoes.
        """
        try:
            df_nova = self.df_main.query("STATUS_CURRENT_PHASE == 'NOVAS SOLICITAÇÕES' and  STATUS_ENVIO_NOVA == 'NAO_ENVIADO'")
            self.lista_novas = [(x['ID_CARDS'], x['TELEFONE'],  f"""Prezado(a) {x['NOME_SOLICITANTE']}, Sua solicitação {x['ID_CARDS']}, foi criada com sucesso. Ela já encontra-se em tratativa e em breve você receberá atualizações. Por favor aguarde. Até breve!""") for y, x in df_nova.iterrows()]
            self.log_info(f'Report_ Relatorio Novas Solicitacoes criado.')
        except Exception as b:
            self.log_erro(f'Report_ {b}')
    
    def relatorio_aprovados(self):
        """
        Criar Lista de Tuplas com (ID_CARD, TELEFONE, MENSAGEM) dos casos aprovados.
        """
        try:
            df_aprovado = self.df_main.query("STATUS_CURRENT_PHASE == 'Agendado ✅' and  STATUS_ENVIO_APROV_REJEI == 'NAO_ENVIADO'")
            self.lista_aprovado = [(x['ID_CARDS'], x['TELEFONE'],  f"""Prezado(a) {x['NOME_SOLICITANTE']}, Conforme solicitação {x['ID_CARDS']}, a manutenção do veículo de placa {x['TEXTO_PLACA']} foi agendada para {x['DATA_AGENDAMENTO']} no fornecedor: {x['NOME_FORNECEDOR']} || {x['ENDERECO_FORNECEDOR']}. A seguinte observação foi deixada: {x['OBSERVACAO_ANALISTA']}\nObrigado!""") for y, x in df_aprovado.iterrows()]
            self.log_info(f'Report_ Relatorio Aprovados criado.')
        except Exception as b:
            self.log_erro(f'Report_ {b}')

    def relatorio_reprovados(self):
        """
        Criar Lista de Tuplas com (ID_CARD, TELEFONE, MENSAGEM) dos casos Reprovados.
        """
        try:
            df_reprovado = self.df_main.query("STATUS_CURRENT_PHASE == 'Reprovado ❌' and  STATUS_ENVIO_APROV_REJEI == 'NAO_ENVIADO'")
            self.lista_reprovado = [(x['ID_CARDS'], x['TELEFONE'],  f"""Prezado(a) {x['NOME_SOLICITANTE']}, Infelizmente a sua solicitação {x['ID_CARDS']} foi REPROVADA pelo seguinte motivo {x['OBSERVACAO_ANALISTA']}. Caso seja necessário, faça uma nova solicitação através do nosso portal: https://portal.pipefy.com/frotaservicos. Obrigado!""") for y, x in df_reprovado.iterrows()]
            self.log_info(f'Report_ Relatorio Reprovado criado.')
        except Exception as b:
            self.log_erro(f'Report_ {b}')

    def relatorio_diario(self):
        """
        Salva o Relatorio com as operações diarias.
        """
        try:
            tempo = str(date.today().strftime("%d-%m-%Y"))
            conn = sqlite3.Connection(f'{self._homedir}My_database\\Dados.db')
            df_diario = pd.read_sql(f"""SELECT * FROM ID_PIPES_DB WHERE DATA_ENVIO = '{tempo}'""", conn)
            df_diario.to_excel(f'{self._homedir}My_docs\\RelatorioDiario{tempo}.xlsx')
            self.log_info(f'Report_ Relatorio diario Salvo.')
        except Exception as b:
            self.log_erro(f'Report_ {b}')

class MessageModel(Report):

    def novas_solicitacoes(self):
        """
        Envio via Whattsapp das novas Solitações.
        """
        try:
            for k in self.lista_novas:
                numero = str(k[1]).replace(' ', '').replace('-', '')
                print(f'chave {k[0]} - telefone{numero} - mensagem {k[2]}')
                py.sendwhatmsg_instantly (numero, k[2], 30, False, 10)
                if rb.locateOnScreen(f'{self._homedir}My_docs\\send_whatss.PNG'):
                    rb.click(rb.center(rb.locateOnScreen(f'{self._homedir}My_docs\\send_whatss.PNG')))
                    os.system('taskkill /f /im chrome.exe')
                self.update_status_envio_nova(k[0])
            self.log_info(f'MessageModel_ {len(self.lista_novas)} enviadas')
        except Exception as b:
            self.log_erro(f'MessageModel_ {b}')

    def aprovadas(self):
        """
        Envio via Whattsapp de Aprovados.
        """
        try:
            for k in self.lista_aprovado:
                numero = str(k[1]).replace(' ', '').replace('-', '')
                print(f'chave {k[0]} - telefone{numero} - mensagem {k[2]}')
                py.sendwhatmsg_instantly (numero, k[2], 30, False, 10)
                if rb.locateOnScreen(f'{self._homedir}My_docs\\send_whatss.PNG'):
                    rb.click(rb.center(rb.locateOnScreen(f'{self._homedir}My_docs\\send_whatss.PNG')))
                    os.system('taskkill /f /im chrome.exe')
                self.update_status_envio_aprov_rejei(k[0])
            self.log_info(f'MessageModel_ {len(self.lista_aprovado)} enviadas')
        except Exception as b:
            self.log_erro(f'MessageModel_ {b}')

    def reprovados(self):
        """
        Envio via Whattsapp de Reprovados.
        """
        try:
            for k in self.lista_reprovado:
                numero = str(k[1]).replace(' ', '').replace('-', '')
                print(f'chave {k[0]} - telefone{numero} - mensagem {k[2]}')
                py.sendwhatmsg_instantly (numero, k[2], 30, False, 10)
                if rb.locateOnScreen(f'{self._homedir}My_docs\\send_whatss.PNG'):
                    rb.click(rb.center(rb.locateOnScreen(f'{self._homedir}My_docs\\send_whatss.PNG')))
                    os.system('taskkill /f /im chrome.exe')
                self.update_status_envio_aprov_rejei(k[0])
            self.log_info(f'MessageModel_ {len(self.lista_reprovado)} enviadas')
        except Exception as b:
            self.log_erro(f'MessageModel_ {b}')          
        

    