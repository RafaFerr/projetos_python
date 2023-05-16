import logging

logging.basicConfig(level=logging.INFO, filename="resources/log/log.txt", format="%(asctime)s / %(levelname)s / %(message)s",
                    datefmt='%d/%m/%Y %I:%M:%S')
from botcity.core import DesktopBot
from botcity.plugins.excel import BotExcelPlugin
from datetime import datetime


class Bot(DesktopBot):
    def action(self, execution=None):

        import pandas as pd
        import numpy as np
        base_partes = BotExcelPlugin()

        base_partes.read(r"C:\RPA\atos_qualificacao\qualificacao\qualificacao\resources\qualificacao_atos_ancelmo.xlsx")
        base_partes.set_active_sheet('qualificacao')
        base_partes._sheets[base_partes.active_sheet].replace(np.nan, '', inplace=True)
        basedados_atos = pd.read_excel(
            r"C:\RPA\atos_qualificacao\qualificacao\qualificacao\resources\qualificacao_atos_ancelmo.xlsx", 'atos',
            keep_default_na=False)
        dados = base_partes.as_list()[1:]

        for i in range(1596):
            print(datetime.now())

            print('0. FOR INICIAL INCREMENTO Nº {}'.format(i))
            logging.info(f'PRIMEIRO ITEM {i}')
            mat_ato = str(basedados_atos["MATRICULA"][i])
            num_protocolo = str(basedados_atos["NUMERO PROTOCOLO"][i])
            num_ato = str(basedados_atos["NUMERO DO ATO"][i])
            # data_ato = (basedados_atos["DATA"][i].strftime("%d/%m/%Y"))
            operacao = str(basedados_atos["OPERACAO"][i])
            ato_completo = str(basedados_atos["ATO_COMPLETO"][i])
            ato_completo2 = str(basedados_atos["ATO_COMPLETO2"][i])
            logging.info(f'MATRICULA {mat_ato} / ATO {num_ato}')

            if not self.find("localiza", matching=0.97, waiting_time=10000):
                self.not_found("localiza")
            self.double_click_relative(102, 45)
            self.delete()
            self.kb_type(text=mat_ato)
            self.enter()
            self.enter()
            print("1. MATRICULA {} SENDO CADASTRADA - ATO {} SENDO CADASTRADA".format(mat_ato, num_ato))
            if not self.find("aba_anotacao", matching=0.97, waiting_time=10000):
                self.not_found("aba_anotacao")
            self.click()
            self.wait(1500)
            self.get_screenshot(
                filepath=rf'C:\RPA\atos_qualificacao\qualificacao\qualificacao\resources\atos\{mat_ato}_ato{num_ato}_original.png')
            self.type_keys(['alt', 'r'])
            # if not self.find("reg_av", matching=0.97, waiting_time=10000):
            #    self.not_found("reg_av")
            # self.click()

            if self.find('ocorrencia', waiting_time=500):
                print('2. OCORRENCIA ENCONTRADA')
                logging.info(f'MATRICULA COM OCORRENCIA')
                self.type_key('0274')
                self.enter()
                self.enter()
                print('3. OCORRENCIA LIBERADA')
            else:
                pass
            self.wait(2000)
            self.delete()
            self.kb_type(text=num_protocolo)
            print("4. NUMERO DE PROTOCOLO {} INSERIDO".format(num_protocolo))
            logging.info(f'PROTOCOLO {num_protocolo}')
            self.enter()
            if self.find('pergunta2', waiting_time=500):
                logging.warning(f'BOTAO pergunta2')
                print('5. PROTOCOLO JÁ ENCERRADO')
                logging.info(f'PROTOCOLO ENCERRADO {num_protocolo}')
                self.type_keys(['alt', 's'])
            elif self.find('atencao2', waiting_time=500):
                logging.warning(f'BOTAO ATENCAO')
                print('6. PROTOCOLO NAO ENCONTRADO')
                logging.info(f'PROTOCOLO {num_protocolo} NAO ENCONTRADO')
                self.enter()
                self.type_keys(['shift', 'home'])
                # if not self.find("campo_protocolo", matching=0.97, waiting_time=1500):
                #    self.not_found("campo_protocolo")
                # self.double_click_relative(183, 43)
                self.delete()
                print('7. LIMPANDO NUMERO DO PROTOCOLO')
                self.enter()
            else:
                pass
            self.wait(500)
            ato = str(basedados_atos["TIPOATO"][i]).upper()
            if ato == "ABERTURA":
                self.type_keys('ab')
            elif ato == "AVERBAÇÃO":
                self.type_keys('av')
            else:
                self.type_key('r')
            self.enter()
            self.kb_type(text=num_ato)
            self.enter()
            if self.find('atencao2', waiting_time=1500):  # PASSANDO POR ESTE IF O ATO JÁ EXISTE CADASTRADO NO SISTEMA
                logging.warning(f'BOTAO ATENCAO')
                logging.info(f'ATO EXISTENTE')
                print('ESTE PRINT É CASO O NUMERO DO ATO EXISTA')
                self.enter()
                self.type_keys(["alt", "c"])
                print('3. CANCELANDO ATO')
                self.wait(500)
                self.type_keys(["alt", "s"])
                self.wait(500)
                self.key_f5()
                print('4. PESQUISAR ATO')
                self.wait(500)
                if not self.find("selecionar_ato", matching=0.99, waiting_time=10000):
                    self.not_found("selecionar_ato")
                self.click_relative(136, 165)
                print('5. SELECIONANDO O PRIMEIRO ATO PARA BUSCA')
                self.wait(500)
                if not self.find("pesquisa_ato", matching=0.97, waiting_time=10000):
                    self.not_found("pesquisa_ato")
                self.double_click_relative(71, 34)
                self.wait(500)
                self.delete()
                print('6. SELECIONANDO CAMPO PARA BUSCA')
                self.copy_to_clipboard(str(basedados_atos["ATO_COMPLETO"][i]))
                self.paste()
                self.enter()
                print('7. BUSCA REALIZADA')

                # ESTE IF IRA ACONTECER CASO ESTEJA CADASTRADO NO REGISTRO UM REGISTRO E NA PLANILHA TENHA UMA AVERBAÇÃO, OU VICE VERSA.
                if self.find('atencao2', waiting_time=1000):
                    logging.warning(f'BOTAO ATENCAO')
                    print('SE PASSOU TEVE TROCA DO TIPO DE ATO')
                    logging.info(f'ATO CADASTRADO INCORRETO - TROCADO {ato_completo2}')
                    self.enter()
                    if not self.find("pesquisa_ato", matching=0.97, waiting_time=10000):
                        self.not_found("pesquisa_ato")
                    self.double_click_relative(71,34)
                    self.wait(500)
                    self.delete()
                    print('7.1 SELECIONANDO CAMPO PARA BUSCA')
                    self.copy_to_clipboard(str(basedados_atos["ATO_COMPLETO2"][i]))
                    self.paste()
                    self.enter()
                    logging.info(f'ATO COMPLETO {ato_completo2}')
                self.key_f6()
                self.wait(1000)
                self.delete()
                print('SE PASSOU DAQUI ATO COMPLETO 1')
                logging.info(f'ATO COMPLETO {ato_completo}')
                self.kb_type(text=num_protocolo)
                print("4. NUMERO DE PROTOCOLO {} INSERIDO".format(num_protocolo))
                self.enter()
                if self.find('pergunta2', waiting_time=500):
                    print('5. PROTOCOLO JÁ ENCERRADO')
                    self.type_keys(['alt', 's'])
                elif self.find('atencao2', waiting_time=500):
                    print('6. PROTOCOLO NAO ENCONTRADO')
                    self.enter()
                    # if not self.find("campo_protocolo", matching=0.97, waiting_time=1500):
                    #    self.not_found("campo_protocolo")
                    # self.double_click_relative(163, 5)
                    self.type_keys(['shift', 'end'])
                    self.delete()
                    print('7. LIMPANDO NUMERO DO PROTOCOLO')
                    self.enter()
                else:
                    pass
                self.wait(500)

                ato = str(basedados_atos["TIPOATO"][i]).upper()
                if ato == "ABERTURA":
                    self.type_keys('ab')
                elif ato == "AVERBAÇÃO":
                    self.type_keys('av')
                else:
                    self.type_key('r')
                self.enter()
                self.kb_type(text=num_ato)
                self.enter()
            self.copy_to_clipboard(basedados_atos["DATA"][i])
            # self.copy_to_clipboard(basedados_atos["DATA"][i].strftime("%d/%m/%Y"))
            self.paste()
            # logging.info(f'DT ATO = {data_ato}')
            self.wait(1000)
            self.enter()
            self.wait(1000)
            self.copy_to_clipboard(basedados_atos["OPERACAO"][i])
            self.paste()
            logging.info(f'OPERACAO = {operacao}')
            self.wait(500)
            self.type_down()
            self.enter()
            self.get_screenshot(
                filepath=rf'C:\RPA\atos_qualificacao\qualificacao\qualificacao\resources\atos\{mat_ato}_ato{num_ato}_partes-originais.png')
            self.wait(500)
            self.key_f4()
            print('8. QUALIFICAR PARTES - ATO NOVO')
            if not self.find("remover_lista", matching=0.97, waiting_time=10000):
                self.not_found("remover_lista")
            self.click()
            while self.find("pergunta2", waiting_time=1000):
                if not self.find("pergunta2", matching=0.97, waiting_time=1000):
                    self.not_found("pergunta2")
                    print('57. REMOVENDO PARTES')
                self.type_keys(['alt', 's'])
                if not self.find("remover_lista", matching=0.97, waiting_time=1000):
                    self.not_found("remover_lista")
                self.click()
            else:
                self.enter()

            for index, cell in enumerate(dados, start=2):
                dados = base_partes.as_list()[1:]
                colunamat = cell[2]
                colunanumato = cell[3]
                colunaqualif = cell[4]
                colunacpf = cell[5]
                colunanome = cell[6]
                colunaestadocivil = cell[7]
                colunanome = str(colunanome)
                colunacpf = str(colunacpf)
                colunaestadocivil = str(colunaestadocivil)
                mat_ato = str(mat_ato)
                colunamat = str(colunamat)
                colunanumato = str(colunanumato)
                print(mat_ato, colunamat)
                print(num_ato, colunanumato)

                def relacionar_parte():
                    self.kb_type(text=colunacpf)
                    if not self.find("campo_nome", matching=0.97, waiting_time=10000):
                        self.not_found("campo_nome")
                    self.double_click_relative(135, 5)
                    self.kb_type(text=colunanome)
                    # self.paste(colunanome)
                    self.enter()
                    if not self.find('codigo_parte',waiting_time=5000):
                        self.wait(3000)
                    else:
                        pass

                    # SE ENTRAR NESTE IF IRÁ CADASTRAR A PARTE NO INDICADOR
                    if self.find('pergunta2', waiting_time=500):
                        print('CRIANDO INDICADOR')
                        logging.info(f'CRIANDO INDICADOR')
                        self.enter()
                        self.wait(2000)
                        self.type_keys(['alt', 's'])
                        if self.find('pergunta2'):
                            self.type_keys(['alt', 's'])
                        self.wait(500)
                    elif self.find('livro5_incluir', waiting_time=3000):
                        print('CRIANDO INDICADOR')
                        logging.info(f'CRIANDO INDICADOR')
                        self.type_keys(['alt', 's'])
                        if self.find('pergunta2'):
                            self.type_keys(['alt', 's'])
                            self.wait(2000)
                    else:
                        pass

                if mat_ato == colunamat and num_ato == colunanumato:
                    print('PLANILHAS VALIDADAS')
                    logging.info(f'PLANILHAS OK')
                    self.key_f5()
                    if not self.find("qualificacao", matching=0.97, waiting_time=10000):
                        self.not_found("qualificacao")
                    self.click_relative(170, 10)
                    self.wait(500)
                    self.page_up()
                    qualificacao = str(colunaqualif)
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Adquirente":
                        self.type_key('ad')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    if qualificacao == "Anuente":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('an')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Arrendatário":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('arr')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Arrendante":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('arrend')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Autor":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('au')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Avalista":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('av')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Cancelado":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('ca')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Cedente":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('ce')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Cessionario":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('ces')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Credor":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('cr')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Custodiante":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('cu')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Devedor":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('de')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Emitente":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('em')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Executado":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('ex')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Exequente":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('exe')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Fiador":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('fi')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Garantidor":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('ga')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Interveniente":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('int')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Interessado":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('in')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Locador":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('lo')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Locatario":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('loc')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Proprietario":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('pr')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Requerente":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('re')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Requerido":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('req')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(
                            f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#
                    # COMEÇA A QUALIFICACAO#
                    if qualificacao == "Transmitente":
                        logging.info(f'QUALIFICACAO {qualificacao} SELECIONADA')
                        self.type_key('tr')
                        self.enter()
                        self.type_key('t')
                        self.enter()
                        self.alt_u()
                        self.enter()
                        logging.info(f'QUALIFICACAO: {qualificacao} / NOME: {colunanome} / CPF: {colunacpf} / EST.CIVIL: {colunaestadocivil}')
                        relacionar_parte()
                    else:
                        pass
                    # TERMINA A QUALIFICACAO#

                    if colunaestadocivil == 'Casado':
                        estado_civil = 'estado_casado'
                    elif colunaestadocivil == 'Divorciado':
                        estado_civil = 'estado_divorcio'
                    elif colunaestadocivil == 'Solteiro':
                        estado_civil = 'estado_solteiro'
                    elif colunaestadocivil == 'Desquitado':
                        estado_civil = 'estado_desquit'
                    elif colunaestadocivil == 'Separado Judicial':
                        estado_civil = 'sep_judicial'
                    elif colunaestadocivil == 'Separado Consensualmente':
                        estado_civil = 'sep_consensual'
                    elif colunaestadocivil == 'Viuvo':
                        estado_civil = 'estado_viuvo'
                    elif colunaestadocivil == 'Casada':
                        estado_civil = 'estado_casado'
                    elif colunaestadocivil == 'Divorciada':
                        estado_civil = 'estado_divorcio'
                    elif colunaestadocivil == 'Solteira':
                        estado_civil = 'estado_solteiro'
                    elif colunaestadocivil == 'Desquitada':
                        estado_civil = 'estado_desquit'
                    elif colunaestadocivil == 'Separada Judicial':
                        estado_civil = 'sep_judicial'
                    elif colunaestadocivil == 'Separada Consensualmente':
                        estado_civil = 'sep_consensual'
                    elif colunaestadocivil == 'Viuva':
                        estado_civil = 'estado_viuvo'
                    elif colunaestadocivil == 'Sem Estado':
                        estado_civil = 'estado_vazio'
                    else:
                        estado_civil = 'em_branco'

                    # ESTE IF IRA VERIFICAR SE A PARTE VERIFICADA É PF OU PJ, SE FOR PJ JÁ IRÁ RELACIONAR A PRIMEIRA PARTE DA LISTA, SE FOR PF, IRÁ VARRER A LISTA ATE ACHAR O ESTADO CIVIL.
                    if not self.find('identifica_pj', waiting_time=2222):
                        logging.info(f"PESSOA FISICA IDENTIFICADA")
                        print('PESSOA FISICA')
                        print(estado_civil)
                        for j in range(9):
                            print('INICIO DO FOR {}'.format(datetime.now()))
                            if self.find('codigo_parte'):
                                print('COM LISTAGEM DE PARTES')
                                if self.find(estado_civil, matching=0.97, waiting_time=2000):
                                    self.wait(500)
                                    self.type_keys(['alt', 'r'])
                                    self.wait(500)
                                    if self.find('atencao2', waiting_time=1000):
                                        logging.warning(f'BOTAO ATENCAO')
                                        self.enter()
                                        print('CLICOU EM ATENÇÃO')
                                    elif self.find('pergunta2', waiting_time=1000):
                                        self.type_keys(['alt', 's'])
                                        print('ALT S NA pergunta2')
                                    else:

                                        pass
                                    print('BREKANDO PARTE EXISTENTE')
                                    break
                                else:
                                    self.type_down()
                                    print(datetime.now())
                            else:
                                print('SEM LISTAGEM')
                                self.type_keys(['alt', 'r'])
                                if self.find('atencao2', waiting_time=1000):
                                    logging.warning(f'BOTAO ATENCAO')
                                    self.enter()
                                    print('CLICOU EM ATENÇÃO')
                                elif self.find('pergunta2', waiting_time=1000):
                                    self.type_keys(['alt', 's'])
                                    print('ALT S NA pergunta2')
                                else:

                                    pass
                                print('BREKANDO PARTE NOVA')
                                break
                        else:  # FIM DO FOR E VAI RELACIONAR O QUE TIVER PRIMEIRO
                            print('FIM DO FOR E VAI RELACIONAR O QUE TIVER PRIMEIRO')
                            self.type_keys(['alt', 'r'])
                            if self.find('atencao2', waiting_time=1000):
                                logging.warning(f'BOTAO ATENCAO')
                                self.enter()
                                print('CLICOU EM ATENÇÃO')
                            elif self.find('pergunta2', waiting_time=1000):
                                self.type_keys(['alt', 's'])
                                print('ALT S NA pergunta2')
                            else:
                                pass
                    else:  # A PARTIR DAQUI FOI ENCONTRADA PESSOA JURIDICA
                        print('ACHOU PJ')
                        self.type_keys(['alt', 'r'])
                        if self.find('atencao2', waiting_time=1000):
                            logging.warning(f'BOTAO ATENCAO')
                            self.enter()
                            print('CLICOU EM ATENÇÃO')
                        elif self.find('pergunta2', waiting_time=1000):
                            self.type_keys(['alt', 's'])
                            print('ALT S NA pergunta2')
                        else:
                            pass
                        print('FIM PJ')

                    self.key_f9()
                    logging.info(f'BOTAO INCLUIR')
                    print('CONCLUIR')
                    base_partes.remove_row(row=1 + 1, sheet='qualificacao')
                    base_partes.write(
                        r"C:\RPA\atos_qualificacao\qualificacao\qualificacao\resources\qualificacao_atos_ancelmo.xlsx")
                    logging.info(f'PARTE {colunanome} REMOVIDA')
                    print('PARTE REMOVIDA')
                    while self.find('pergunta2', waiting_time=1000):
                        self.type_keys(['alt', 's'])
                        print('CONFIRMANDO NACIONALIDADE E/OU INDISPONIBILIDADE')
                        logging.info(f'CONFIRMANDO NACIONALIDADE OU INDISPONIBILIDADE')
                    else:
                        pass
                else:  # OS DADOS DAS PLANILHAS NAO BATEM
                    #logging.info(f'OS DADOS DAS PLANILHAS NAO BATEM')
                    pass
            self.key_esc()
            logging.info(f'QUALIFICACAO CONCLUIDA')
            self.wait(2000)
            if self.find('gerar_doi', waiting_time=2000):
                self.type_keys(['alt','g'])
            else:
                pass
            self.get_screenshot(
                filepath=rf'C:\RPA\atos_qualificacao\qualificacao\qualificacao\resources\atos\{mat_ato}_ato{num_ato}_partes-cadastradas.png')
            self.wait(500)
            self.type_keys(['alt', 's'])
            logging.info(f'SALVANDO ATO')
            while self.find('pergunta2', waiting_time=1000):
                self.type_keys(['alt', 's'])
                print('CONFIRMANDO NACIONALIDADE E/OU INDISPONIBILIDADE')
            else:
                pass
            self.get_screenshot(
                filepath=rf'C:\RPA\atos_qualificacao\qualificacao\qualificacao\resources\atos\{mat_ato}_ato{num_ato}_alterado.png')
            self.wait(2500)

            if not self.find("tarefa", matching=0.97, waiting_time=1000):
                self.not_found("tarefa")
            self.click()
            print("72. CLICK 1 TAREFA")
            self.wait(500)
            self.click()
            print("73. CLICK 2 TAREFA")
            self.wait(500)
            if not self.find("tarefa2", matching=0.97, waiting_time=1000):
                self.not_found("tarefa2")
            self.click()
            print("74. CLICK TAREFA 2")
            self.wait(500)
            if not self.find("aba_tabela2", matching=0.97, waiting_time=1000):
                self.not_found("aba_tabela2")
            self.click()
            print("75. CLICK 1 ABA TABELA")

    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()
