import pywhatkit
import pandas as pd
from datetime import datetime, timedelta

caminho_planilha = r'C:\caminho\para\sua\planilha.xlsx'
df = pd.read_excel(caminho_planilha)

ja_enviado = df['Enviado'].tolist()

lista_nomes = df['Nome'].tolist()
print(lista_nomes)


lista_contatos = df['Telefone'].tolist()
for i in range(len(lista_contatos)):
    lista_contatos[i] = '+55' + str(lista_contatos[i])
print (lista_contatos)


hora_atual = datetime.now().hour
minuto_atual = datetime.now().minute

hora_envio = hora_atual
minuto_envio = minuto_atual + 1


if minuto_envio >= 60:
    minuto_envio -= 60
    hora_envio += 1

for i in range (len(lista_contatos)):
    # Verifique se a mensagem já foi enviada para este contato.
    if df.loc[i, 'Enviado'] == 'S':
        # Converta a data de envio para um objeto datetime.
        data_envio = df.loc[i, 'Data de Envio'].to_pydatetime()

        # Verifique se faz mais de 30 dias desde que a mensagem foi enviada.
        if datetime.now() - data_envio > timedelta(days=30):
            # Aqui você pode adicionar o código para enviar uma nova mensagem.
            # mensagem enviada caso o cliente ja tenha o seu produto
            nova_mensagem = f'Olá {lista_nomes[i]}, tudo bem? Só estou passando para perguntar se está precisando de um reabastecimento de produtos. Se precisar, estou à disposição!'
            pywhatkit.sendwhatmsg(lista_contatos[i], nova_mensagem, hora_envio, minuto_envio)

            # Atualize a data de envio.
            df.loc[i, 'Data de Envio'] = datetime.now().strftime('%d/%m/%Y')
        continue
    
    
    mensagem = f'''(Sua mensagem de vendas)'''
    pywhatkit.sendwhatmsg(lista_contatos[i], mensagem, hora_envio, minuto_envio)
    minuto_envio += 1
    if minuto_envio == 60:
        minuto_envio = 0
        hora_envio += 1
    # Marque este contato como "enviado".
    df.loc[i, 'Enviado'] = 'S'
    df.loc[i, 'Data de Envio'] = datetime.now().strftime('%d/%m/%Y')

# Salve o DataFrame de volta na planilha.
df.to_excel(r'C:\caminho\para\sua\planilha.xlsx', index=False)
