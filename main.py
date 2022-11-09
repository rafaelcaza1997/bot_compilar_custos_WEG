import os
import pandas as pd
from datetime import datetime
from pathlib import Path
import win32com.client as win32
import configparser
import argparse
import requests
from datetime import datetime


cfg = configparser.ConfigParser()


def get_dolar(data):
    data_obj = datetime.strptime(data,"%d/%m/%Y")
    data_conver = datetime.strftime(data_obj, "%m-%d-%Y")
    try:
        requisicao = requests.get("https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/CotacaoDolarDia(dataCotacao=@dataCotacao)?@dataCotacao='"+ str(data_conver) +"'&$top=1&$format=json&$select=cotacaoCompra")
        cotacao = requisicao.json()
        dolar = cotacao['value'][0]['cotacaoCompra']
        return float(dolar)
    except:
        return float(0)


def checar_path(path):
    while os.path.exists(path) == False:
        os.makedirs(path)
        if os.path.exists(path) == True:
            break
    return path


 
def ler_arquivos(filename):
    arquivos_antigos = []
    try:
        with open(filename) as f:
            for linha in f:
                arquivos_antigos.append(linha.replace("\n",""))
    except:
        pass

    return arquivos_antigos


def verificar_centro(str):
    for key in centros:
        if centros[key]["prefixo"] in str:
            return key
    return "Erro"


def converter_data(str, centro):
    data = str
    for value in centros[centro].values():
        data = data.replace(value,"")
    data = data.strip()
    return datetime.strptime(data, "%d.%m.%Y")


def converter_num(x):
    try:
        x = float(x.strip().replace("/t","").replace(".","").replace(",","."))
    except:
        pass
    return x


def converter_tam_lote(x):
    if '-' not in str(x):
        x = str(x.replace(".",""))
    return x.strip()


def get_paths_csv(limite = 20):
    pathlist = Path(path_pasta_origem).glob('**/*.txt')
    
    dict_arquivo = {}
    for path in pathlist:

        centro = verificar_centro(path.name.strip())
        data = converter_data(path.name.strip(), centro)

        if data not in dict_arquivo.keys():
            dict_arquivo[data] = {
            centro:path,
            }
        else:
            dict_arquivo[data][centro] = path

    dict_filtrado = {}
    for item in sorted(dict_arquivo.items()):
        dict_filtrado[item[0]] = item[1]

    lista_keys = [key for key in dict_filtrado.keys()]
    lista_keys.reverse()
    lista_keys = lista_keys[0:limite]
    lista_keys.reverse()

    dict_limitado = {}
    for key in lista_keys:
        dict_limitado[key] = dict_filtrado[key]

    return dict_limitado


def create_historico_custos(dict_values):

    df_hist_last = pd.DataFrame(columns=['procv_material'])
    df_lote_last = pd.DataFrame(columns=['procv_material'])

    df_resumo_last = pd.DataFrame(columns=['procv_material','Material','Centro','UMAv'])


    df_descricao = pd.DataFrame(columns=['Material','Descricao'])

    ultima_data = None
    penultima_data = None

    vetor_datas = []


    for data in dict_values:
        data_str =  datetime.strftime(data,"%d/%m/%Y")

        vetor_datas.append(data_str)

        df_historico_temp = pd.DataFrame()
        df_lote_temp = pd.DataFrame()
        df_resumo_temp = pd.DataFrame()

        penultima_data = ultima_data
        ultima_data = data_str

        for centro in dict_values[data]:

            df = pd.read_csv(dict_values[data][centro], sep="|",header=0, encoding='latin-1', dtype = str)

            dict_primeira_conv = {}
            for col in df.columns:
                dict_primeira_conv[col] = col.strip()
            df = df.rename(columns=dict_primeira_conv)
            df = df.rename(columns=convrt_cabecalho)

            df_descricao = pd.merge(df_descricao, df[['Material', 'Descricao']], how = "outer", on = ['Material', 'Descricao'])
            df_descricao.drop_duplicates(subset = "Material", inplace = True)

            try:
                df_custos = df[['Material', 'Centro', "Tam_lote", "UMAv", "Custo"]]
            except:
                convrt_cabecalho_2 = {'Tam.lote cÃ¡lc.csts.':'Tam_lote'}
                df = df.rename(columns=convrt_cabecalho_2)
                df_custos = df[['Material', 'Centro', "Tam_lote", "UMAv", "Custo"]]

            df_custos['Material'] = df_custos['Material'].apply(lambda x: str(x).strip())
            df_custos['Custo'] = df_custos['Custo'].apply(lambda x: converter_num(x))
            df_custos = df_custos.rename(columns={'Custo': data_str})
            df_custos['Tam_lote'] = df_custos['Tam_lote'].apply(lambda x: converter_tam_lote(x))
            df_custos['UMAv'] = df_custos['UMAv'].apply(lambda x: str(x).strip())
            df_custos['Centro'] = df_custos['Centro'].apply(lambda x: str(int(str(x).strip())))
            df_custos['procv_material'] = df_custos['Material']+df_custos['Centro']


            df_resumo_temp = pd.concat([df_resumo_temp, df_custos[['procv_material','Material','Centro','UMAv']]])

            df_historico_temp = pd.concat([df_historico_temp, df_custos[['procv_material',data_str]]])

            df_custos.drop(columns = data_str, index = 1, inplace = True)
            df_custos = df_custos.rename(columns={'Tam_lote': data_str})
            df_lote_temp = pd.concat([df_lote_temp, df_custos[['procv_material',data_str]]])


        df_historico = pd.merge(df_hist_last, df_historico_temp[['procv_material', data_str]], how = "outer", on = ['procv_material'])
        df_lote = pd.merge(df_lote_last, df_lote_temp[['procv_material', data_str]], how = "outer", on = ['procv_material'])


        df_resumo = pd.merge(df_resumo_last, df_resumo_temp[['procv_material','Material','Centro','UMAv']], how = "outer", on = ['procv_material','Material','Centro','UMAv'])

        df_hist_last = df_historico
        df_lote_last = df_lote
        df_resumo_last = df_resumo


    df_resumo = pd.merge(df_resumo, df_descricao, how = "left", on = ['Material'])

    df_resumo = pd.merge(df_resumo, df_historico[['procv_material', penultima_data, ultima_data]], how = "left", on = ['procv_material'])

    dict_convrt = {
        penultima_data: "Custo " + penultima_data,
        ultima_data: "Custo " + ultima_data,
    }

    df_resumo.rename(columns=dict_convrt, inplace = True)

    # df_resumo = pd.merge(df_resumo, df_descricao, how = "left", on = ['Material'])
    df_resumo["Reajuste"] = df_resumo[dict_convrt[ultima_data]] - df_resumo[dict_convrt[penultima_data]]
    df_resumo["Reajuste_%"] = df_resumo["Reajuste"] / df_resumo[dict_convrt[penultima_data]]


    df_resumo = pd.merge(df_resumo, df_lote[["procv_material", df_lote.columns[-1]]], how = "left", on = ['procv_material'])
    df_resumo.rename(columns={df_lote.columns[-1] : 'Tam_lote'}, inplace = True)


    # df_resumo = df_resumo[['procv_material', 'Material', 'Descricao', 'Centro', 'UMAv',	dict_convrt[penultima_data], dict_convrt[ultima_data], 'Reajuste', 'Reajuste_%']]

    df_resumo = df_resumo[['procv_material', 'Material', 'Descricao', 'Centro', 'Tam_lote', 'UMAv',	dict_convrt[penultima_data], dict_convrt[ultima_data], 'Reajuste', 'Reajuste_%']]
        
    dict_convrt = {
        "Custo " + penultima_data : "custo_antigo",
        "Custo " + ultima_data : "custo_atual",
    }

    df_resumo_padronizado = df_resumo.rename(columns=dict_convrt)

    df_resumo.sort_values( by = 'Reajuste_%', ascending = False, inplace = True, ignore_index = True)

    df_qntd = pd.read_excel(path_arquivo_quantidade, header = 3)

    df_filtrado_qntd = pd.DataFrame()

    df_filtrado_qntd["Material"] = df_qntd["Material"].apply(lambda x: str(x.replace("#/","")))
    df_filtrado_qntd["Volume"] = df_qntd["Unnamed: 10"]

    df_quantidade_total = df_filtrado_qntd.groupby(by = 'Material', as_index = False).sum()

    df_resumo = df_resumo.merge(df_quantidade_total, how='left', on='Material')
    # df_resumo.fillna(0, inplace = True)

    df_parametros = pd.DataFrame(data = {"Custo anterior": ["Custo " + penultima_data], "Custo atual": ["Custo " + ultima_data]})


    def get_df(path, idioma):
        df = pd.read_csv(path, sep="|",header=1, encoding='latin-1', dtype = str)
        df.drop(index= 0, inplace = True) # Conferir axis
        df.drop(columns = ["Unnamed: 0","Unnamed: 4","MTyp","TpMt"], inplace = True, errors = 'ignore')
        df.columns = ['Material','Descricao_' + idioma]
        df["Material"] = df["Material"].apply(lambda x: str(x).strip())
        df.drop( index = df.loc[(df["Material"] == "Material")].index, inplace = True) 
        return df



    df_descricao_ES = get_df(path_descricao_ES, "ES")
    df_descricao_US = get_df(path_descricao_US, "US")

    df_merge = df_descricao_ES.merge(df_descricao_US, how = "left", on = "Material").dropna()
    df_merge.reset_index(drop = True, inplace = True)

    # df_descricao_PT = df_resumo_padronizado[["Material", "Descricao"]]
    df_resumo_padronizado = df_resumo_padronizado.merge(df_merge, how = "left", on = "Material")

    dict_row = {   
    'procv_material' : ['dolar'],
    }

    for data in vetor_datas:
        dict_row[data] = [get_dolar(data)] 

    df2 = pd.DataFrame.from_dict(dict_row)
    df_historico = pd.concat([df_historico, df2])
    df_historico.reset_index(drop = True, inplace = True)

    return df_historico, df_lote, df_resumo, df_resumo_padronizado, df_parametros


def enviar_email(path_arquivo_custos, destinatarios_email):

# criar a integração com o outlook
    outlook = win32.Dispatch('outlook.application')

    # criar um email
    email = outlook.CreateItem(0)

    # configurar as informações do seu e-mail
    col_datas[0]
    col_datas[1]

    email.To = destinatarios_email
    email.Subject = f"Planilha Análise de custos - {str(col_datas[0])} - {str(col_datas[1])}"
    email.HTMLBody = f"""
    <p>Planilha de comparação de custos entre os dias {str(col_datas[0])} - {str(col_datas[1])}.</p>
    <p>E-mail enviado automaticamente referente a atualização de custos.</p>
    """

    anexo = path_arquivo_custos
    email.Attachments.Add(anexo)

    email.Send()
    print("Email Enviado")




if __name__ == '__main__':
    
    parser = argparse.ArgumentParser(description='path_workspace argumento para definir o path da pasta onde o arquivo está localizado') 
    parser.add_argument('--path_workspace') 
        
    args = parser.parse_args()
    path_workspace = args.path_workspace

    if path_workspace != None:
        path_padrao = path_workspace.replace("\\","//") 
    else:
        path_padrao = os.getcwd().replace("\\","//") 

    path_inicial = path_padrao+'//cfg.ini'


    template_inicial = [
    "[path]",
    "pasta_destino_arquivo_planilha_preco = "+path_padrao,
    "pasta_destino_arquivo_comparacao_custos = "+path_padrao,
    "pasta_origem_arquivos_custos = "+path_padrao+"//csv_custos",
    "path_arquivo_quantidade = "+path_padrao,
    "path_descricao_ES = "+ path_padrao + "//ES.txt",
    "path_descricao_US = "+ path_padrao + "//EN.txt",
    "[email]",
    "enviar_email = True",
    "destinatarios = email1@gmail.com;email2@gmail.com",
    "",
    "[formatacao_arquivos_txt]",
    "centro_1609_prefixo = ",
    "centro_1609_sufixo =  ",
    "centro_1607_prefixo = ",
    "centro_1607_sufixo = ",
    "formato_data = &d.&m.&Y",
    ]


    if not os.path.isfile(path_inicial):
        with open(path_inicial, 'w') as f:
            for linha in template_inicial:
                f.write(linha+"\n")

    cfg.read(path_inicial)

    path_destino_planilha_preco = checar_path(cfg['path']['pasta_destino_arquivo_planilha_preco'].replace("\\","//"))
    path_destino_analise_custos = checar_path(cfg['path']['pasta_destino_arquivo_comparacao_custos'].replace("\\","//"))
    path_pasta_origem = checar_path(cfg['path']['pasta_origem_arquivos_custos'].replace("\\","//"))

    path_arquivo_quantidade = cfg['path']['path_arquivo_quantidade'].replace("\\","//")
    
    path_descricao_ES = cfg['path']['path_descricao_ES'].replace("\\","//")
    path_descricao_US = cfg['path']['path_descricao_US'].replace("\\","//")


    bool_email = cfg["email"]["enviar_email"]
    if bool_email == 'True':
        gatilho_enviar_email = True
    else:
        gatilho_enviar_email = False

    destinatarios_email = cfg["email"]["destinatarios"]

    if destinatarios_email == "":
        gatilho_enviar_email = False


    centro_1609_prefixo = cfg["formatacao_arquivos_txt"]["centro_1609_prefixo"]
    centro_1609_sufixo = cfg["formatacao_arquivos_txt"]["centro_1609_sufixo"]
    centro_1607_prefixo = cfg["formatacao_arquivos_txt"]["centro_1607_prefixo"]
    centro_1607_sufixo = cfg["formatacao_arquivos_txt"]["centro_1607_sufixo"]

    centros = {
        "1609":{
            "prefixo" : centro_1609_prefixo.strip(),
            "sufixo" : centro_1609_sufixo.strip(),
        },
        "1607":{
            "prefixo" : centro_1607_prefixo.strip(),
            "sufixo" : centro_1607_sufixo.strip(),
        },
        }



    formato_data = cfg["formatacao_arquivos_txt"]["formato_data"].replace("&","%")
    path_arquivos_lidos = path_padrao+"//banco_dados_arquivos.txt"

    atualizar_planilha = False

    convrt_cabecalho = {'Material':'Material',
                        'Texto breve de material':'Descricao',
                        'Centro':'Centro',
                        'Total Un.':'Custo',
                        'Tam.lote cálc.csts.':'Tam_lote',
                        'UMAv':'UMAv',
                        'Ano':'Ano',
                        'Per':'Per'}

    #Rotina Principal
    arquivos_antigos = ler_arquivos(path_arquivos_lidos)
    pathlist = Path(path_pasta_origem).glob('**/*.txt')

    print("Path Arquivos csv")
    for path in pathlist:
        print(path.name)
        if path.name not in arquivos_antigos:
            atualizar_planilha = True
            arquivos_antigos.append(path.name)

    print("Resumo Paths")
    print("Path Padrão: ",path_padrao)
    print("Path .ini: ",path_inicial)
    print("Pasta origem: ",path_pasta_origem)


    print("Atualizar planilha? ",atualizar_planilha)


    if atualizar_planilha:
        data_hoje = datetime.strftime(datetime.now(),"%d_%m_%Y")

        path_analise_custos = path_destino_analise_custos + "//Comparacao_custo_"+data_hoje+".xlsx"
        print("Atualizando arquivos")

        dict = get_paths_csv( limite = 20)
        df_historico, df_lote, df_resumo, df_resumo_padronizado, df_parametros = create_historico_custos(dict)

        col_datas = [str(col) for col in df_historico.columns[-2::]]

        writer = pd.ExcelWriter(path_destino_planilha_preco+"//tabela_custos.xlsx", engine='xlsxwriter')
        df_resumo_padronizado.to_excel(writer, sheet_name='custos', index = False)
        df_lote.to_excel(writer, sheet_name='historico_lote', index = False)
        df_parametros.to_excel(writer, sheet_name='parametros', index = False)
        df_historico.to_excel(writer, sheet_name='historico_custos', index = False)
        writer.save()

        writer = pd.ExcelWriter(path_analise_custos, engine='xlsxwriter')
        df_resumo.to_excel(writer, sheet_name='custos', index = False)
        df_historico.to_excel(writer, sheet_name='historico_custos', index = False)
        writer.save()


        # rotina_atualizar_arquivos()
        if gatilho_enviar_email:
            enviar_email(path_analise_custos, destinatarios_email)
    else:
        print("Arquivos já atualizados")

    with open(path_arquivos_lidos, 'w') as f:
        for path in arquivos_antigos:
            f.write(path+"\n")

