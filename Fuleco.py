import requests
from io import BytesIO
import PySimpleGUI as sg
import winsound
import pyautogui
import pandas as pd
import time

def buscar_entrega(data_inicial, data_final, entrega):
    url = "https://api-dw.bseller.com.br/webquery/execute-excel/QRY0166"
    headers = {
        "Content-Type": "application/json",
        "X-Auth-Token": "746E4AA1FE7A85BCE053A7F3A8C0AAED"
    }
    body = {
        "parametros": {
            "P_DT_INICIAL": data_inicial,
            "P_DT_FINAL": data_final,
            "P_ID_PLANTA": "MECAJ", 
            "P_SIT": None,
            "P_STATUS_OM": None,
            "P_ID_ENTREGA": entrega
        }
    }
    try:
        response = requests.post(url, headers=headers, json=body)
        if response.status_code == 200:
            excel_data = response.content
            df_primario = pd.read_excel(BytesIO(excel_data))
            df_primario = df_primario.rename(columns={'Id item': 'Item WMS', 'Qtde. Separada': 'Quantidade'})
            df_primario['Entrega'] = df_primario['Entrega'].astype(str)
            df_primario['Item WMS'] = df_primario['Item WMS'].astype(str)
            df_primario['Descrição item'] = df_primario['Descrição item'].astype(str)


            df_primario['Ean'] = df_primario['Ean'].astype(str)

            #df_primario['Ean'] = df_primario['Ean'].apply(
                #lambda x: '0' + x if 11 <= len(x) <= 12 else x
            #)

            df_primario['Ean'] = df_primario['Ean'].apply(
            lambda x: x.zfill(13) if len(x) < 13 else x
            )

            df_pedido = df_primario.groupby(['Item WMS', 'Ean']).agg({
                'Entrega': 'first',
                'Quantidade': 'sum',
                'Descrição item' : 'first'
            }).reset_index()
            return df_pedido
        else:
            sg.popup_error("Falha ao carregar os dados. Verifique os parâmetros e tente novamente.")
            return None
    except Exception as e:
        sg.popup_error(f"Ocorreu um erro durante a solicitação: {str(e)}")
        return None

def combinar_ean():
    url = "https://api-dw.bseller.com.br/webquery/execute-excel/SIGEQ233"
    headers = {
        "Content-Type": "application/json",
        "X-Auth-Token": "746E4AA1FE7A85BCE053A7F3A8C0AAED"
    }
    body = {
        "parametros": {
            "P_ID_CIA": 11228,
            "P_ID_DEPART": None,
            "P_ID_FAMI": None,
            "P_ID_SETOR": None,
            "P_ID_SITUACAO": None
        }
    }

    response = requests.post(url, headers=headers, json=body)
    if response.status_code == 200:
        excel_data = response.content
        maisa_ean = pd.read_excel(BytesIO(excel_data))
        maisa_ean = maisa_ean.rename(columns={'Item': 'Id item'})
        dados = maisa_ean.groupby('Id item').agg({
            'Ean': lambda x: ','.join(map(str, x)),
            'Descrição': 'first'
        }).reset_index()
        return dados
    else:
        sg.popup_error("Falha ao carregar os dados da segunda API.")
        return None

def main(data_inicial, data_final, entrega):
    df_entrega = buscar_entrega(data_inicial, data_final, entrega)
    if df_entrega is None:
        return None

    df_ean = combinar_ean()
    if df_ean is None:
        return None

    df_entrega_expanded = df_entrega.assign(Ean=df_entrega['Ean'].str.split(',')).explode('Ean')
    df_ean_expanded = df_ean.assign(Ean=df_ean['Ean'].str.split(',')).explode('Ean')

    df_final = pd.merge(df_entrega_expanded, df_ean_expanded, on='Ean', how='outer')[['Ean', 'Quantidade', 'Descrição', 'Item WMS']]

    # Preencher Item WMS com base na descrição se houver valores faltantes
    df_final['Item WMS'] = df_final.groupby('Descrição')['Item WMS'].transform('first').fillna(df_final['Item WMS'])

    return df_final

def verificar_faltas_sobras(codigos_escaneados, df):
    faltas = {}
    sobras = {}
    escaneados_por_item = {item: 0 for item in df["Item WMS"].unique()}

    for codigo in codigos_escaneados:
        if codigo in df["Ean"].tolist():
            item = df[df["Ean"] == codigo]["Item WMS"].values[0]
            escaneados_por_item[item] += 1

    for item in escaneados_por_item.keys():
        quantidade_pedido = df[df["Item WMS"] == item]["Quantidade"].sum()
        quantidade_escaneada = escaneados_por_item[item]
        if quantidade_escaneada < quantidade_pedido:
            faltas[item] = quantidade_pedido - quantidade_escaneada
        elif quantidade_escaneada > quantidade_pedido:
            sobras[item] = quantidade_escaneada - quantidade_pedido

    return faltas, sobras

def alerta():
    winsound.PlaySound("alert.wav", winsound.SND_FILENAME)

sg.theme('Dark blue')

layout = [
    #[sg.Image(filename='Lojado.png', size=(500, 100))],

    [sg.Text('Data Inicial'), sg.InputText(key='-DATA_INICIAL-')],
    [sg.Text('Data Final'), sg.InputText(key='-DATA_FINAL-')],
    [sg.Text('ID Entrega'), sg.InputText(key='-ID_ENTREGA-')],
    [sg.Button('Buscar')],
    [sg.Button('Limpar')],
    [sg.Text('Insira Ean:', size=(16, 1)), sg.InputText(key='-CODIGO-', enable_events=True, disabled=True)],

    #[sg.Button('Enter', key='-ENTER-')],
    
    [sg.Text('Faltas:', size=(10, 1)), sg.Output(key='-FALTAS-', size=(50, 5))],
    [sg.Text('Sobras:', size=(10, 1)), sg.Output(key='-SOBRAS-', size=(50, 5))],
]

window = sg.Window('Fuleco V.1' , layout)

codigos_escaneados = []
df_pedido = None

def processar_codigo(codigo):
    if len(codigo) in [13] and codigo != '':
        print("Código escaneado:", codigo)
        if codigo not in df_pedido["Ean"].tolist():
            sg.popup_no_buttons('Item não faz parte do pedido!', title='Alerta', auto_close=True, auto_close_duration=3)
            alerta()
            window['-CODIGO-'].update('')
        else:
            codigos_escaneados.append(codigo)
            faltas, sobras = verificar_faltas_sobras(codigos_escaneados, df_pedido)

            faltas_str = '\n'.join([f'Item WMS: {k}, Faltando: {v}' for k, v in faltas.items()])
            sobras_str = '\n'.join([f'Item WMS: {k}, Sobrando: {v}' for k, v in sobras.items()])
            window['-FALTAS-'].update(faltas_str)
            window['-SOBRAS-'].update(sobras_str)

            # Remove one leading zero for external system, if it exists
            if codigo.startswith('0'):
                codigo_modificado = codigo[1:]
            else:
                codigo_modificado = codigo

            window['-CODIGO-'].Widget.selection_range(0, 'end')
            window['-CODIGO-'].Widget.event_generate('<<Copy>>')

            pyautogui.click(x=487, y=375)
            pyautogui.typewrite(codigo_modificado)  # Usar o código modificado sem um zero à esquerda
            pyautogui.press('enter')
            pyautogui.click(x=1042, y=366)
            pyautogui.press('enter')

            posicao_atual = pyautogui.position()
            pyautogui.moveTo(posicao_atual)
            window['-CODIGO-'].Widget.focus_set()

            window['-CODIGO-'].update('')
            pyautogui.press('enter')

            # Verificação e ação caso não haja faltas nem sobras
            if not faltas and not sobras:
                pyautogui.click(x=645, y=619)
                pyautogui.press('enter')

                # Reiniciar dados para nova validação de pedido
                codigos_escaneados.clear()
                window['-FALTAS-'].update('')
                window['-SOBRAS-'].update('')
                window['-CODIGO-'].update('')
                sg.popup('Pedido validado com sucesso! Pronto para nova validação.', title='Sucesso')

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED:
        break
    elif event == 'Buscar':
        df_pedido = main(values['-DATA_INICIAL-'], values['-DATA_FINAL-'], values['-ID_ENTREGA-'])
        if df_pedido is not None:
            sg.popup('Dados carregados com sucesso!', title='Sucesso')
            window['-CODIGO-'].update(disabled=False)
    elif event == 'Limpar':
        codigos_escaneados.clear()
        df_pedido = None
        window['-DATA_INICIAL-'].update('')
        window['-DATA_FINAL-'].update('')
        window['-ID_ENTREGA-'].update('')
        window['-CODIGO-'].update('')
        window['-CODIGO-'].update(disabled=True)
        window['-FALTAS-'].update('')
        window['-SOBRAS-'].update('')
        sg.popup('Dados limpos com sucesso!', title='Sucesso')
    elif event == '-ENTER-' or event == '-CODIGO-' and len(values['-CODIGO-']) in [13]:
        processar_codigo(values['-CODIGO-'])

    
window.close()
