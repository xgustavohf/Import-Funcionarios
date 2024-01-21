from tkinter import Tk, filedialog, messagebox
import customtkinter as ctk
import threading
import time
import shutil
import traceback
import os
import datetime
import cx_Oracle
from cryptography.fernet import Fernet
import logging


caminho_instant_client = r'/opt/oracle/instantclient_11_2'

os.environ['PATH'] = f'{caminho_instant_client};{os.environ["PATH"]}'

logging.basicConfig(level=logging.INFO) 

root = ctk.CTk()
root.geometry("570x400")
root.title("Integração de Funcionários Grupo Brandili Têxtil")
root.resizable(False, False)
root._set_appearance_mode("dark") 
ctk.set_default_color_theme("blue")

arquivos_path = "Arquivos"


def verificar_conexao():
    server = input_server.get()
    user = input_user.get()
    password = input_password.get()
    base = input_base.get()
    port = input_port.get()

    string_conexao = f'{user}/{password}@{server}:{port}/{base}'

    try:
        conn = cx_Oracle.connect(string_conexao)
        conn.close()
        messagebox.showinfo("Conexão", "Conexão bem sucedida!")
        logging.info("Conexão bem sucedida ao banco de dados.")
    except cx_Oracle.DatabaseError as e:
        messagebox.showerror("Conexão", f"Falha na conexão: {str(e)}")
        logging.error(f"Falha na conexão ao banco de dados: {str(e)}")


try:
    with open('configuracao/chave.key', 'rb') as key_file:
        chave = key_file.read()
except FileNotFoundError:
    chave = Fernet.generate_key()
    with open('configuracao/chave.key', 'wb') as key_file:
        key_file.write(chave)

cipher = Fernet(chave)


def buscar_pasta():
    pasta = filedialog.askdirectory(initialdir=input_arquivos.get())
    input_arquivos.delete(0, 'end')
    input_arquivos.insert(0, pasta)

def carregar_caminho():
    try:
        with open('configuracao/caminho.txt', 'r') as f:
            caminho_salvo = f.read().strip()
            return caminho_salvo
    except FileNotFoundError:
        return ""

def salvar_caminho():
    caminho = input_arquivos.get()
    with open('configuracao/caminho.txt', 'w') as f:
        f.write(caminho)

def carregar_dados():
    try:
        with open('configuracao/dados.txt', 'rb') as f:
            dados_criptografados = f.read()
            dados_decifrados = cipher.decrypt(dados_criptografados).decode('utf-8')

            linhas = dados_decifrados.split('\n')
            server = linhas[0].split(': ')[1].strip()
            user = linhas[1].split(': ')[1].strip()
            password = linhas[2].split(': ')[1].strip()
            base = linhas[3].split(': ')[1].strip()
            port = linhas[4].split(': ')[1].strip()

            input_server.delete(0, 'end')
            input_user.delete(0, 'end')
            input_password.delete(0, 'end')
            input_base.delete(0, 'end')
            input_port.delete(0, 'end')

            input_server.insert(0, server)
            input_user.insert(0, user)
            input_password.insert(0, password)
            input_base.insert(0, base)
            input_port.insert(0, port)

    except FileNotFoundError:
        pass

def salvar_dados():
    server = input_server.get()
    user = input_user.get()
    password = input_password.get()
    base = input_base.get()
    port = input_port.get()

    dados = f'Servidor: {server}\nUsuário: {user}\nSenha: {password}\nBanco: {base}\nPorta: {port}'

    dados_criptografados = cipher.encrypt(dados.encode('utf-8'))

    os.makedirs('configuracao', exist_ok=True)

    with open('configuracao/dados.txt', 'wb') as f:
        f.write(dados_criptografados)

def salvar_conexao():
    salvar_dados()
    verificar_conexao()
    salvar_caminho()


label_title = ctk.CTkLabel(root, text="Processo de Importação de Funcionários", text_color="white", bg_color="#242424", fg_color="transparent", font=("Roboto Bold", 19))
label_title.place(x=120, y=20)

label_fita_head = ctk.CTkFrame(root, width=580, height=1, bg_color="white")
label_fita_head.place(x=0, y=70)

label_arquivo = ctk.CTkLabel(root, text="Local Arquivos:", text_color="white", bg_color="#242424", fg_color="transparent", font=("Roboto", 13))
label_arquivo.place(x=20, y=100)

input_arquivos = ctk.CTkEntry(root, width=350, bg_color="#242424", placeholder_text="Escolha o caminho do arquivo que deseja importar  ➡️", font=("Roboto", 12))
input_arquivos.place(x=115, y=100)

caminho_inicial = carregar_caminho()
input_arquivos.insert(0, caminho_inicial) 

button_arquivos = ctk.CTkButton(root, text="Localizar", text_color="white", width=50, bg_color="#242424", hover_color="#0eda25", command=buscar_pasta)
button_arquivos.place(x=470, y=100)

label_logs = ctk.CTkLabel(root, text="Logs:", text_color="white", bg_color="#242424", fg_color="transparent", font=("Roboto", 12))
label_logs.place(x=75, y=150)

input_logs = ctk.CTkEntry(root, width=350, bg_color="#242424", placeholder_text="Escolha o caminho onde você deseja salvar os logs  ➡️", font=("Roboto", 12))
input_logs.place(x=115, y=150)

button_logs = ctk.CTkButton(root, text="Localizar", text_color="white", width=50, bg_color="#242424", hover_color="#0eda25")
button_logs.place(x=470, y=150)

label_fita_body = ctk.CTkFrame(root, width=580, height=1, bg_color="white")
label_fita_body.place(x=0, y=210)

label_server = ctk.CTkLabel(root, text="Host", text_color="white", bg_color="#242424", fg_color="transparent", font=("Roboto", 12))
label_server.place(x=20, y=230)

input_server = ctk.CTkEntry(root, bg_color="#242424", width=98, font=("Roboto", 12))
input_server.place(x=20, y=255)

label_user = ctk.CTkLabel(root, text="Usuário", text_color="white", bg_color="#242424", fg_color="transparent", font=("Roboto", 12))
label_user.place(x=140, y=230)

input_user = ctk.CTkEntry(root, bg_color="#242424", width=80, font=("Roboto", 12))
input_user.place(x=140, y=255)

label_password = ctk.CTkLabel(root, text="Senha", text_color="white", bg_color="#242424", fg_color="transparent", font=("Roboto", 12))
label_password.place(x=243, y=230)

input_password = ctk.CTkEntry(root, bg_color="#242424", width=80, show="*", font=("Roboto", 12))
input_password.place(x=243, y=255)

label_base = ctk.CTkLabel(root, text="Serviço", text_color="white", bg_color="#242424", fg_color="transparent", font=("Roboto", 12))
label_base.place(x=420, y=230)

input_base = ctk.CTkEntry(root, bg_color="#242424", width=100, font=("Roboto", 12))
input_base.place(x=420, y=255)

label_port = ctk.CTkLabel(root, text="Porta", text_color="white", bg_color="#242424", fg_color="transparent", font=("Roboto", 12))
label_port.place(x=345, y=230)

input_port = ctk.CTkEntry(root, bg_color="#242424", width=50, font=("Roboto", 12))
input_port.place(x=345, y=255)

btn_verific = ctk.CTkButton(root, text="OK", text_color="white", width=20, bg_color="#242424", hover_color="#0eda25", command=salvar_conexao, font=("Roboto", 12))
btn_verific.place(x=525, y=255)


def auto_executar():
    btn_execute.invoke()
    root.after(15000, root.destroy)

def gerar_log(arquivo, mensagem, sucesso=True):
    log_folder = os.path.join(os.path.abspath(os.path.dirname(__file__)), "Logs", datetime.datetime.now().strftime("%Y-%m-%d"))

    try:
        os.makedirs(log_folder, exist_ok=True)
    except Exception as e:
        logging.error(f"Erro ao criar pasta de logs: {str(e)}")

    status = "SUCESSO" if sucesso else "ERRO"
    log_file_path = os.path.join(log_folder, f"log_{datetime.datetime.now().strftime('%H-%M-%S')}_{status}_{arquivo}.txt")

    with open(log_file_path, 'w') as log_file:
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_file.write(f"{timestamp} - {status} - Arquivo: {arquivo} - {mensagem}\n")


def converter_data(data_brasileira):
    if data_brasileira == '00/00/0000':
        return None 
    else:
        
        data_obj = datetime.datetime.strptime(data_brasileira, '%d/%m/%Y')
        data_formatada = data_obj.strftime('%Y-%m-%d')
        return data_formatada

def executar_importacao():
    def importacao_thread():
        user = input_user.get()
        password = input_password.get()
        server = input_server.get()
        port = input_port.get()
        base = input_base.get()

        string_conexao = f'{user}/{password}@{server}:{port}/{base}'

        try:
            conn = cx_Oracle.connect(string_conexao)
            cursor = conn.cursor()

            pasta = input_arquivos.get()
            registros_importados_total = 0

            log_folder = os.path.join(os.path.abspath(os.path.dirname(__file__)), "Logs",
                                      datetime.datetime.now().strftime("%Y-%m-%d"))

            try:
                os.makedirs(log_folder, exist_ok=True)
            except Exception as e:
                logging.error(f"Erro ao criar pasta de logs: {str(e)}")

            log_file_path = os.path.join(log_folder, "importacao_log.txt")
            logging.basicConfig(filename=log_file_path, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


            for arquivo in os.listdir(pasta):
                if arquivo.endswith(".txt"):
                    arquivo_path = os.path.join(pasta, arquivo)
                    
                    try:
                        with open(arquivo_path, 'r') as file:
                            registros_importados = 0
                            for linha in file:
                                dados = linha.strip().split('|')
                                cod_func = dados[0].strip()
                                ind_situacao = dados[1].strip()
                                nom_func = dados[2].strip()
                                nro_cnpj = dados[3].strip()
                                ind_pontualidade = dados[4].strip()
                                ind_bloqueio = dados[5].strip()
                                dsc_endereco = dados[6].strip()
                                ind_uf = dados[7].strip()
                                ind_pf_pj = dados[8].strip()
                                nro_rg = dados[9].strip()
                                numcpf = dados[10].strip()
                                nom_cidade = dados[11].strip()
                                nro_cep = dados[12].strip()
                                nro_telefone = dados[13].strip()
                                desconhecido = dados[14].strip()
                                ind_sem_credito = dados[15].strip()
                                dat_nascimento = converter_data(dados[16].strip())
                                dat_admissao = converter_data(dados[17].strip())
                                nro_ddd = dados[18].strip()
                                ind_sexo = dados[19].strip()
                                dsc_email = dados[20].strip()
                                nom_bairro = dados[21].strip()
                                vlr_limite_cred = dados[22].strip()
                                dsc_logradouro = dados[23].strip()
                                nom_pais = dados[24].strip()
                                nro_logradouro = dados[25].strip()
                                ind_estado_civil = dados[26].strip()


                                cursor.execute("SELECT * FROM VOL_FUNC_CAJOVIL WHERE COD_FUNC = :1", (cod_func,))
                                existing_record = cursor.fetchone()

                                if existing_record:
                                    cursor.execute("UPDATE VOL_FUNC_CAJOVIL SET COD_FUNC=:1, IND_SITUACAO=:2, NOM_FUNC=:3, NRO_CNPJ=:4, IND_PONTUALIDADE=:5, IND_BLOQUEIO=:6, DSC_ENDERECO=:7, IND_UF=:8, IND_PF_PJ=:9, NRO_RG=:10, NUMCPF=:11, NOM_CIDADE=:12, NRO_CEP=:13, NRO_TELEFONE=:14, DESCONHECIDO=:15, IND_SEM_CREDITO=:16, DAT_NASCIMENTO=:17, DAT_ADMISSAO=:18, NRO_DDD=:19, IND_SEXO=:20, DSC_EMAIL=:21, NOM_BAIRRO=:22, VLR_LIMITE_CRED=:23, DSC_LOGRADOURO=:24, NOM_PAIS=:25, NRO_LOGRADOURO=:26, IND_ESTADO_CIVIL=:27 WHERE COD_FUNC=:1", 
                                    (cod_func, ind_situacao, nom_func, nro_cnpj, ind_pontualidade, ind_bloqueio, dsc_endereco, ind_uf, ind_pf_pj, nro_rg, numcpf, nom_cidade, nro_cep, nro_telefone, desconhecido, ind_sem_credito, dat_nascimento, dat_admissao, nro_ddd, ind_sexo, dsc_email, nom_bairro, vlr_limite_cred, dsc_logradouro, nom_pais, nro_logradouro, ind_estado_civil))
                                else:
                                    cursor.execute("INSERT INTO VOL_FUNC_CAJOVIL (COD_FUNC, IND_SITUACAO, NOM_FUNC, NRO_CNPJ, IND_PONTUALIDADE, IND_BLOQUEIO, DSC_ENDERECO, IND_UF, IND_PF_PJ, NRO_RG, NUMCPF, NOM_CIDADE, NRO_CEP, NRO_TELEFONE, DESCONHECIDO, IND_SEM_CREDITO, DAT_NASCIMENTO, DAT_ADMISSAO, NRO_DDD, IND_SEXO, DSC_EMAIL, NOM_BAIRRO, VLR_LIMITE_CRED, DSC_LOGRADOURO, NOM_PAIS, NRO_LOGRADOURO, IND_ESTADO_CIVIL) VALUES (:1, :2, :3, :4, :5, :6, :7, :8, :9, :10, :11, :12, :13, :14, :15, :16, :17, :18, :19, :20, :21, :22, :23, :24, :25, :26, :27)",
                                    (cod_func, ind_situacao, nom_func, nro_cnpj, ind_pontualidade, ind_bloqueio, dsc_endereco, ind_uf, ind_pf_pj, nro_rg, numcpf, nom_cidade, nro_cep, nro_telefone, desconhecido, ind_sem_credito, dat_nascimento, dat_admissao, nro_ddd, ind_sexo, dsc_email, nom_bairro, vlr_limite_cred, dsc_logradouro, nom_pais, nro_logradouro, ind_estado_civil))
                                
                                registros_importados += 1 

                            registros_importados_total += registros_importados

                            gerar_log(arquivo, f"Arquivo importado com sucesso. Linhas importadas: {registros_importados}")


                    except Exception as e:
                        gerar_log(arquivo, f"Erro na importação: {str(e)}", sucesso=False)
                        traceback.print_exc()

                        log_file_path = os.path.join(log_folder, f"log_{datetime.datetime.now().strftime('%H-%M-%S')}_ERRO_{arquivo}.txt")

                        with open(log_file_path, 'w') as log_file:
                            log_file.write(f"Erro na Importação do arquivo {arquivo}: {str(e)}\n")


            conn.commit()
            conn.close()

            logging.info(f"Importação concluída com sucesso! {registros_importados_total} linhas foram importadas.")

            if registros_importados_total > 0:
                try:
                    processed_folder = os.path.join(os.path.abspath(os.path.dirname(__file__)), "Arquivos",
                                                    "Arquivos Processados",
                                                    datetime.datetime.now().strftime("%Y-%m-%d"))
                    os.makedirs(processed_folder, exist_ok=True)

                    shutil.move(arquivo_path, os.path.join(processed_folder, arquivo))

                except Exception as move_error:
                    logging.error(f"Erro ao mover arquivo processado: {str(move_error)}")

            btn_execute.configure(state="normal")

        except cx_Oracle.DatabaseError as e:
            logging.error(f"Falha na Conexão: {str(e)}")
            traceback.print_exc()
            messagebox.showerror("Conexão", f"Falha na conexão: {str(e)}")

        btn_execute.configure(state="normal")

    btn_execute.configure(state="disabled")
    threading.Thread(target=importacao_thread).start()

btn_execute = ctk.CTkButton(root, text="Executar", text_color="white", width=200, bg_color="#242424",
                            hover_color="#0eda25", font=("Roboto", 15), command=executar_importacao)
btn_execute.place(x=180, y=350)

label_fita_rodape = ctk.CTkFrame(root, width=580, height=1, bg_color="white")
label_fita_rodape.place(x=0, y=320)

carregar_dados()
root.after(10000, auto_executar)
root.mainloop()