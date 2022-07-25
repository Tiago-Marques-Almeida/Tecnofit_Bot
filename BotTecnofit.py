import os
import shutil
import time
from datetime import date, datetime, timedelta
import configTecnofit
import unidecode

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementClickInterceptedException

import pandas as pd
from bs4 import BeautifulSoup


class Bot():

    def __init__(self):
        self.data = datetime.now()
        self.reset_ambiente()
        self.driver = self.get_driver()
        self.list_unidade = []
        #self.data = datetime.now()
        self.dia_anterior = (self.data - timedelta (1)).strftime('%d-%m-%Y')
        self.data_31_dias = (self.data + timedelta(31)).strftime('%d-%m-%Y')
        self.hora_tempo_real = self.data.strftime('%H:%M')
        self.data_inicio = self.data.strftime('%d/%m/%Y')
        self.data_pasta = self.data.strftime('%d-%m-%Y')
        print(self.data_pasta)
        self.data_inicio_historico = self.data.strftime('01/%m/%Y')
        
        self.data_fim = self.data.strftime('%d/%m/%Y')
        
        self.run()


    def get_driver(self):
        chrome_options = Options()
        #chrome_options.add_argument('--headless')
        chrome_options.add_experimental_option("prefs", {"download.default_directory": configTecnofit.DIRETORIO_ARQUIVOS_TEMP})
        s=Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=s,options=chrome_options)
        driver.maximize_window()
        return driver

    def cria_diretorio(self, caminho):
        if not os.path.isdir(caminho):
            os.mkdir(caminho)

    def reset_ambiente(self):
        list_caminhos = [
            configTecnofit.DIRETORIO_ARQUIVOS_TEMP
                   ]
        for caminho in list_caminhos:
            if os.path.isdir(caminho):
                shutil.rmtree(caminho)
            self.cria_diretorio(caminho)

    def login(self, usuario, senha):
        self.driver.get(configTecnofit.URL_SISTEMA)
        self.retorna_elemento('XPATH', '/html/body/app-root/app-login/div/div/div[1]/div/div[2]/div/form/div[1]/div[1]/input').send_keys(usuario)
        self.retorna_elemento('XPATH', '/html/body/app-root/app-login/div/div/div[1]/div/div[2]/div/form/div[1]/div[2]/input').send_keys(senha)
        self.driver.find_element(By.XPATH,'/html/body/app-root/app-login/div/div/div[1]/div/div[2]/div/form/div[2]/div/div/button[2]').submit()
        #tempo necessário pois o carregamento do site demora
        time.sleep(1)

    def logoff(self):
        self.driver.get('https://sis.bluefitacademia.com.br/ControlGym/View/dashboard-homolog.php')
        self.driver.find_element(By.XPATH,'//*[@id="mainwrapper"]/div[1]/div[2]/ul/li[3]/div/div/ul/li[3]/a').click()

    def aguarda_download(self):
        seconds = 1
        time.sleep(seconds)
        dl_wait = True
        while dl_wait and seconds < 60:
            time.sleep(2)
            dl_wait = False
            for fname in os.listdir(configTecnofit.DIRETORIO_ARQUIVOS_TEMP):
                if fname.endswith('.crdownload'):
                    dl_wait = True
            seconds += 1
        return seconds

    def renomar_arquivo(self, original, alterar):
        #self.convert_to_parquet(original)
        self.aguarda_download()
        os.rename(original, alterar)

    def convert_to_parquet(self, caminho, alterar):
        self.aguarda_download()
        if(os.path.isfile(caminho)):
            if(caminho.find('.xls')>=0):
                if(alterar.find('Vendas_Realizadas')>=0):
                    soup = BeautifulSoup(open(caminho, 'r', encoding="utf-8"), 'lxml')
                    tabela = soup.find_all('table')
                    pd.read_html(str(tabela).replace('colspan',''), encoding='utf8', decimal=',', thousands='.')[0].to_parquet(alterar)
                else:
                    pd.read_html(open(caminho, 'r', encoding='utf8'), encoding='utf8', decimal=',', thousands='.')[0].to_parquet(alterar)
            os.remove(caminho)

    def cria_diretorio(self, caminho):
        if not os.path.isdir(caminho):
            os.mkdir(caminho)

    def retorna_elemento(self, funcao, path):
        self.aguardar_elemento(funcao, path)
        return self.driver.find_element(getattr(By,funcao), path)

    def aguardar_elemento(self, funcao, path):
        WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((getattr(By,funcao), path)))

    def remover_notificacao(self):
        time.sleep(3)
        elements = self.driver.find_elements(By.TAG_NAME, 'tecnofit-micro-notification')
        for e in elements:
            self.driver.execute_script("""
            var element = arguments[0];
            element.parentNode.removeChild(element);
            """, e)


    def Ativos_dia_a_dia(self):
        print('Ativos_dia_a_dia')
        caminho_inicio = configTecnofit.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xls'
        caminho_fim = configTecnofit.DIRETORIO_ARQUIVOS + '\\' + self.data_pasta + '\\' + 'Ativos_dia' + '-' + self.nome_da_unidade   + '.parquet'

        if os.path.isfile(caminho_fim):
            return 0

        #Href página ativos dia a dia
        time.sleep(2)
        self.driver.get('https://app.tecnofit.com.br/relatorio/ativosDia')
        self.remover_notificacao()
        
        self.retorna_elemento('XPATH', '//*[@id="frmPesquisa"]/div/div[2]/div/div/div[3]/div/div/div/label/div/ins').click()

        di = self.retorna_elemento('ID', 'data_inicial')#.setAttribute("value", data_inicio)
        df = self.retorna_elemento('ID', 'data_final')#.setAttribute("value", data_inicio)

        self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",di, self.data_inicio_historico[3:])
        self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",df, self.data_fim[3:])

        #ativa a caixa "mostrar ativos dia a dia"

        #CLica em pesquisar e apresenta os dados
        self.retorna_elemento('ID', 'btnPesquisa').click()
        #clica para exportar a tabela e aguarda o download

        self.verifica_block()
        #clicar para baixar o relatório
        self.aguarda_download()

        if(len(self.driver.find_elements(By.ID, 'btnExporar'))):
            self.retorna_elemento('ID','btnExporar').click()
            self.convert_to_parquet(caminho_inicio, caminho_fim)
        else:
            print(f'Não há registros para Ativos dia para a unidade {self.nome_da_unidade} no periodo entre {self.data_inicio_historico[3:]} e {self.data_fim[3:]}')


    def verifica_block(self):
        contador = 0 
        time.sleep(1)
        while(len(self.driver.find_elements(By.CLASS_NAME, 'block-spinner-bar'))):
            time.sleep(2)
            if(contador == 10):
                return 0
            contador+=1

    def clientes_bloqueados(self):
        print('clientes_bloqueados')
        
        caminho_inicio = configTecnofit.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xls'
        caminho_fim = configTecnofit.DIRETORIO_ARQUIVOS + '\\' + self.data_pasta + '\\' + 'Clientes_Bloqueados' + '-' + self.nome_da_unidade   + '.parquet'
        if os.path.isfile(caminho_fim):
            return 0

        time.sleep(2)
        #abre pagina com relatório de bloqueados
        self.driver.get('https://app.tecnofit.com.br/relatorio/clientesBloqueados')
        self.remover_notificacao()
        
        '''#coloca para hoje
            #clica na data
        di = self.retorna_elemento('ID', 'data_inicial')#.setAttribute("value", data_inicio)
        df = self.retorna_elemento('ID', 'data_final')#.setAttribute("value", data_inicio)

        self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",di, self.data_inicio_historico)
        self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",df, self.dia_anterior)'''

        #Escreve bloqueio de 90 dias atrás
        self.retorna_elemento('XPATH', '//*[@id="frmPesquisa"]/div/div[2]/div/div/div[1]/div/div/input[2]').send_keys('90')
        time.sleep(1)
        
        #clica em selecionar motivo de bloqueio
        self.retorna_elemento('XPATH', '/html/body/div[2]/div[2]/div[1]/div/div/form/div/div[2]/div/div/div[2]/div/div/a/span[1]').click()
        #seleciona pagamento em atraso
        #self.retorna_elemento('ID', 'select2-chosen-1').click()
        self.retorna_elemento('ID', 's2id_autogen1_search').send_keys('atraso', Keys.ENTER)
        self.retorna_elemento('NAME', 'dias_de_bloqueio_fim').send_keys('atraso', Keys.ENTER)
        
        #Pesquisar
        self.retorna_elemento('ID', 'btnPesquisa').click()
        self.verifica_block()
        #clicar para baixar o relatório
        self.aguarda_download()

        if(len(self.driver.find_elements(By.ID, 'btnExporar'))):
            self.retorna_elemento('ID','btnExporar').click()
            self.convert_to_parquet(caminho_inicio, caminho_fim)
        else:
            print(f'Não há registros para Clientes Bloqueados para a unidade {self.nome_da_unidade} no periodo entre {self.data_inicio_historico} e {self.dia_anterior}')


    def contratos_cancelados(self):
        print('contratos_cancelados')
        
        caminho_inicio = configTecnofit.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xls'
        caminho_fim = configTecnofit.DIRETORIO_ARQUIVOS + '\\' + self.data_pasta + '\\' + 'Contratos_Cancelados' + '-' + self.nome_da_unidade   + '.parquet'
        
        if os.path.isfile(caminho_fim):
            return 0
            
        time.sleep(1)
        #site contratos cancelados
        self.driver.get('https://app.tecnofit.com.br/relatorio/contratoCancelado')

        self.remover_notificacao()
        #Seleciona data

        di = self.retorna_elemento('ID', 'data_inicial')#.setAttribute("value", data_inicio)
        df = self.retorna_elemento('ID', 'data_final')#.setAttribute("value", data_inicio)

        self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",di, self.data_inicio_historico)
        self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",df, self.dia_anterior)

        #clicar no botão pesquisar
        self.retorna_elemento('ID', 's2id_cancellationType').click()
        self.retorna_elemento('ID', 's2id_autogen3_search').send_keys('cancelados', Keys.ENTER)
        self.retorna_elemento('XPATH', '//*[@id="btnPesquisa"]').click()
        self.verifica_block()
        #clicar para baixar o relatório
        self.aguarda_download()
        
        if(len(self.driver.find_elements(By.ID, 'btnExporar'))):
            self.retorna_elemento('ID','btnExporar').click()
            self.convert_to_parquet(caminho_inicio, caminho_fim)
        else:
            print(f'Não há registros para Contratos Cancelados para a unidade {self.nome_da_unidade} no periodo entre {self.data_inicio_historico} e {self.dia_anterior}')


    def vendas_realizadas(self):
        print('vendas_realizadas')
        time.sleep(2)
        
    
        caminho_inicio = configTecnofit.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xls'
        caminho_fim = configTecnofit.DIRETORIO_ARQUIVOS + '\\' + self.data_pasta + '\\' + 'Vendas_Realizadas' + '-' + self.nome_da_unidade   + '.parquet'
    
        if (os.path.isfile(caminho_fim)):
            return 0

        #acessa site
        self.driver.get('https://app.tecnofit.com.br/relatorio/vendaRealizada')
        self.remover_notificacao()
        #self.retorna_elemento('TAG_NAME', 'body').send_keys(Keys.ESCAPE, Keys.ESCAPE, Keys.ESCAPE)
        #opção para baixar histórico
        
        di = self.retorna_elemento('ID', 'data_inicial')#.setAttribute("value", data_inicio)
        df = self.retorna_elemento('ID', 'data_final')#.setAttribute("value", data_inicio)

        self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",di, self.data_inicio_historico)
        self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",df, self.dia_anterior)
        self.verifica_block()
        #marca caixa cpf aluno
        self.retorna_elemento('XPATH', '//*[@id="frmPesquisa"]/div/div[2]/div/div[2]/div[4]/div/div/div/label/div/ins').click()
        #clica em pesquisar
        self.retorna_elemento('ID', 'btnPesquisa').click()
        
        self.verifica_block()
        
        #clica no botão para ver todas as possibilidades
        caixaop = self.driver.find_elements(By.XPATH, '//*[@id="grid"]/div/div[2]/div/div/div')
        
        if(len(caixaop)>0):
            caixaop[0].click()
            #seleciona todos por página
            all_options = self.retorna_elemento('XPATH', '//*[@id="grid"]/div/div[2]/div/div/div/ul')
            for i in all_options.find_elements(By.TAG_NAME, 'li'):
                if(i.text == 'Todos'):
                    i.click()
                    break
        self.verifica_block()
        self.aguarda_download()

        try:
            self.retorna_elemento('ID', 'frmPesquisa').click()
        except:
            print("Erro no em tentar clicar no formulario")
        if(len(self.driver.find_elements(By.ID, 'btnExportToExcel'))):
            self.retorna_elemento('ID','btnExportToExcel').click()
            self.convert_to_parquet(caminho_inicio, caminho_fim)
        else:
            print(f'Não há registros para Vendas Realizadas para a unidade {self.nome_da_unidade} no periodo entre {self.data_inicio_historico} e {self.dia_anterior}')


    def recorrencias(self):
        print('recorrencias')
        
        caminho_inicio = configTecnofit.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xls'
        caminho_fim = configTecnofit.DIRETORIO_ARQUIVOS + '\\' + self.data_pasta + '\\' + 'Recorrencias' + '-' + self.nome_da_unidade   + '.parquet'
        
        if os.path.isfile(caminho_fim):
            return 0
        
        time.sleep(2)
        #acessando site de recorrencia
        self.driver.get('https://app.tecnofit.com.br/relatorio/pagamentoRecorrente')
        #remover notificação
        self.remover_notificacao()
        
        di = self.retorna_elemento('ID', 'data_inicial')#.setAttribute("value", data_inicio)
        df = self.retorna_elemento('ID', 'data_final')#.setAttribute("value", data_inicio)

        self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",di, self.data_inicio_historico)
        self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",df, self.dia_anterior)
        
        #selecionando pendente

        self.retorna_elemento('ID', 'select2-chosen-2').click()
        self.retorna_elemento('ID', 's2id_autogen2_search').send_keys('Pendente', Keys.ENTER)
        self.verifica_block()
       
        self.retorna_elemento('ID', 'btnPesquisa').click()
        self.verifica_block()
        #clicar para baixar o relatório
        
        if(len(self.driver.find_elements(By.ID, 'btnExporar'))):
            self.retorna_elemento('ID','btnExporar').click()
            self.convert_to_parquet(caminho_inicio, caminho_fim)
        else:
            print(f'Não há registros para Recorrencias para a unidade {self.nome_da_unidade} no periodo entre {self.data_inicio_historico} e {self.dia_anterior}')


    def get_cliente_ativo(self):
        print('get_cliente_ativo')
        time.sleep(5)
        self.driver.get('https://app.tecnofit.com.br/ng/dashboard/financeiro')
        self.aguarda_loanding()
        time.sleep(4)
        self.remover_notificacao()
        filtrado = list(filter(lambda x : x.text.strip()=='Configurar painel', self.driver.find_elements(By.TAG_NAME, 'h2')))
        if(len(filtrado)):
            self.retorna_elemento('XPATH', '/html/body/app-root/app-dashboard-configuration/div/app-suggestion/div/div[2]/div[1]/div/ul/li[3]').click()
            list(filter(lambda x : x.text.strip()=='Gerar painel', self.driver.find_elements(By.TAG_NAME, 'button')))[0].click()
        else:
            titulos = self.retorna_elemento('CLASS_NAME','small-cards').find_elements(By.TAG_NAME, 'app-card-active-customers')
            for i in titulos:
                if(i.get_attribute('cardid')=='18'):
                    quantidade_aluno = i.find_element(By.CLASS_NAME, 'card-body').find_element(By.CLASS_NAME, 'ng-star-inserted').text.strip()
                    
                    caminho_fim = configTecnofit.DIRETORIO_ARQUIVOS + '\\' + self.data_pasta + '\\' + 'Alunos_Ativo_Tempo_Real' + '-' + self.nome_da_unidade   + '.parquet'
                    
                    if os.path.isfile(caminho_fim):
                        os.remove(caminho_fim)
                    
                    pd.DataFrame({
                        "DATA":[self.data.strftime('%d/%m/%Y')],
                        "HORA": [self.hora_tempo_real],
                        "UNIDADE": [self.nome_da_unidade],
                        "QTD_ALUNO": [quantidade_aluno]
                    }).to_parquet(caminho_fim)
                    
    def contratos_cancelados_agendados(self):
        print('contratos_cancelados_agendados')
        
        caminho_inicio = configTecnofit.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xls'
        caminho_fim = configTecnofit.DIRETORIO_ARQUIVOS + '\\' + self.data_pasta + '\\' + 'Contratos_Cancelados_agendamento' + '-' + self.nome_da_unidade   + '.parquet'
        
        if (os.path.isfile(caminho_fim)):
            return 0
        
        time.sleep(1)
        #site contratos cancelados
        self.driver.get('https://app.tecnofit.com.br/relatorio/contratoCancelado')

        self.remover_notificacao()
        #Seleciona data

        di = self.retorna_elemento('ID', 'data_inicial')#.setAttribute("value", data_inicio)
        df = self.retorna_elemento('ID', 'data_final')#.setAttribute("value", data_inicio)

        self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",di, self.data_pasta)
        self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",df, self.data_31_dias)
       
        #clicar no botão pesquisar
        self.retorna_elemento('ID', 's2id_cancellationType').click()
        self.retorna_elemento('ID', 's2id_autogen3_search').send_keys('Agendamento', Keys.ENTER)
        self.retorna_elemento('XPATH', '//*[@id="btnPesquisa"]').click()
        self.verifica_block()
        #clicar para baixar o relatório
        self.aguarda_download()
        
        if(len(self.driver.find_elements(By.ID, 'btnExporar'))):
            self.retorna_elemento('ID','btnExporar').click()
            self.convert_to_parquet(caminho_inicio, caminho_fim)
        else:
            print(f'Não há registros para Contratos Cancelados para a unidade {self.nome_da_unidade} no periodo entre {self.data_pasta} e {self.data_31_dias}')


    def vendas_por_tipo_item(self):
        print('vendas_por_tipo_item')
        
        caminho_inicio = configTecnofit.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xls'
        caminho_fim = configTecnofit.DIRETORIO_ARQUIVOS + '\\' + self.data_pasta + '\\' + 'venda_tipo_item' + '-' + self.nome_da_unidade   + '.parquet'
        
        if (os.path.isfile(caminho_fim)):
            return 0
        
        time.sleep(1)

        #Abre site de vendas tipo
        self.driver.get('https://app.tecnofit.com.br/relatorio/vendasTipo')

        self.remover_notificacao()
        time.sleep(1)

        #ler arquivo de_para e puxa os planos a serem baixados
        d_unidade_plano = self.ler_de_para()

        #lista otodos os planos possível
        itens = self.retorna_elemento('XPATH', '//*[@id="campo_contrato"]/div/div/div/div/ul')

        #abre o leque de opcoes de planos
        try:
            self.retorna_elemento('XPATH', '/html/body/div[2]/div[2]/div[1]/div/div/form/div/div[2]/div/div/div[3]/div/div/div/button').click()
        except TimeoutException:
            print('processando Academia Evolve')
            self.retorna_elemento('XPATH', '//*[@id="campo_contrato"]/div/div/div/button/span[1]').click()

        #convertando nome unidade - tirando
        unidade_v = self.nome_da_unidade
        unidade_sa = unidecode.unidecode(unidade_v)

        #seleciona planos a serem filtrados
        for i in d_unidade_plano[unidade_sa]:            
            print('Selecionando Planos...')
            for iten in itens.find_elements(By.TAG_NAME, 'a'):
                if(i == iten.text):
                    iten.click()                  
        time.sleep(3)
        
        #clica para mostrar apenas novas vendas
        self.retorna_elemento('XPATH', '//*[@id="mostrar-apenas-novas-vendas"]/div/div/div/div/label/div/ins').click()
        
        #Clica em pesquisar
        try:
            self.retorna_elemento('ID', 'btnPesquisa').click()
            time.sleep(5)
        except ElementClickInterceptedException:
            print('Tentando achar o botão.')
            self.retorna_elemento('XPATH', '//*[@id="campo_contrato"]/div/div/div/button/span[1]').click()
            self.retorna_elemento('ID', 'btnPesquisa').click()            
            time.sleep(5)


        #clica para baixar relatório
        if(len(self.driver.find_elements(By.ID, 'btnExporar'))):
            self.retorna_elemento('ID','btnExporar').click()
            self.convert_to_parquet(caminho_inicio, caminho_fim)
        else:
            print(f'Não há registros para Contratos Cancelados para a unidade {self.nome_da_unidade} no periodo entre {self.data_inicio} e {self.data_fim}')


   
    def ler_de_para(self):

        #ler arquivo de_para
        df = pd.read_excel(configTecnofit.DIRETORIO_ARQUIVOS_DE_PARA +  '\\de_para_vendas_consultores_planos.xlsx')
        d_unidade_plano = {}
        for unidade, plano in zip(df['Unidade'], df['plano']):
            if(unidade in d_unidade_plano.keys()):
                d_unidade_plano[unidade].append(plano)
            else:
                d_unidade_plano.update({unidade:[plano]})
        return d_unidade_plano 


        
    def aguarda_loanding(self):
        contador = 0
        while(len(self.driver.find_elements(By.CLASS_NAME, 'overlay-loading'))):
            time.sleep(3)
            if contador==10:
                return 0
            contador+=1

    def listar_unidade(self):
        time.sleep(2)
        self.driver.get('https://app.tecnofit.com.br/cadastro/empresa/selecionarEmpresa')

        caixa3 = self.retorna_elemento('ID', 'divListaEmpresa')
        unidades = caixa3.find_elements(By.TAG_NAME, 'a')
        
        list_unidade = []
        ignorar = ''
        
        for unidade in unidades:
            if (unidade.text == ignorar):
                print('contrato cancelado')
            else:
                list_unidade.append(unidade.text)
            
        return list_unidade

    def mudar_unidade(self):
        time.sleep(3)
        self.driver.get('https://app.tecnofit.com.br/cadastro/empresa/selecionarEmpresa')

        self.retorna_elemento('ID', 'pesq').send_keys(self.list_unidade.pop(), Keys.ENTER)
        
        if(len(self.driver.find_elements(By.CLASS_NAME, 'blockUI')) >0):
            time.sleep(3)

        caixa3 = self.retorna_elemento('ID', 'divListaEmpresa')
        unidade = caixa3.find_elements(By.TAG_NAME, 'a')[0]
        
        unidade.click()
        self.nome_da_unidade = unidade.text
        print(self.nome_da_unidade)
        return self.listar_unidade
        

    def run(self):
        for credencial in configTecnofit.CREDENCIAIS:
            self.login(credencial['usuario'], credencial['senha'])
            self.cria_diretorio(configTecnofit.DIRETORIO_ARQUIVOS + '\\' + self.data_pasta)
            
            self.list_unidade = self.listar_unidade()
            
            while(len(self.list_unidade)):
                self.mudar_unidade()
                self.vendas_realizadas()
                self.Ativos_dia_a_dia()
                self.clientes_bloqueados()
                self.contratos_cancelados_agendados()
                self.contratos_cancelados()
                self.recorrencias()
                self.get_cliente_ativo()
                self.vendas_por_tipo_item()                
        self.driver.close()
        self.driver.quit()

if __name__ == '__main__':
    tentativas = 0
    while (tentativas<1):
        try:
            bot = Bot()
            tentativas = 1
        except Exception as e:
            print(e)
            tentativas += 1
            if (tentativas >=1):
                print('criando arquivo de log...')
                with open (r'C:\robo\Tecnofit-bot\log\log.txt', 'w') as arq:
                    arq.write(str(e) + datetime.today())
            
        time.sleep(30)
