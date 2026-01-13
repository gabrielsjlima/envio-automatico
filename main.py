import win32com.client as client
import datetime as dt
import pandas as pd

#carregando a planilha de cobranca e informando a data de hoje
tabela = pd.read_excel("cobranca.xlsx")
hoje = dt.datetime.now()

#login com o Outlook e vinculando o e-mail que quero que mande
outlook = client.Dispatch("Outlook.Application")
emissor = outlook.session.Accounts["financeiro@iberreta.com.br"]

#pega todos os dados da tabela
dados = tabela[["EMAIL", "VALOR", "VENCIMENTO"]].values.tolist()

#o for percorre aos dados coletados
for dado in dados:
    destinatario = dado[0]
    valor = dado[1]
    vencimento = dado[2]

    #formata a data de vencimento
    vencimento = vencimento.strftime("%d/%m/%Y")

    #comandos para enviar o e-mail
    mensagem = outlook.CreateItem(0)
    mensagem.To = destinatario
    mensagem.Subject = "Aviso de vencimento"
    mensagem.HTMLBody = f"""
    <p>Prezado Cliente,</p>

    <p>Estamos entrando em contato para informar que há um boleto com vencimento em <strong>{vencimento}</strong> no valor de <strong>R${valor:.2f}</strong></p>
    
    <p>Por gentileza, pedimos que verifique e caso não tenha o boleto, favor entrar em contato conosco!</p>

    <p>Caso já tenha efetuado o pagamento, por favor desconsidere este aviso.</p>

    <p>Att,</p>

    <p>Gabriel Lima <br>
    Irmãos Berreta LTDA</p>

    """
    mensagem._oleobj_.Invoke(*(64209,0,8,0,emissor))

    mensagem.Save()
    mensagem.Send()

    print("E-mail enviado com sucesso!")