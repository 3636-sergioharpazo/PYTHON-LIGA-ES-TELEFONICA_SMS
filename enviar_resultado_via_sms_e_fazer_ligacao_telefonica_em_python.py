from ctypes.wintypes import MSG
from re import I
import pandas as pd
from twilio.rest import Client
from twilio.twiml.voice_response import VoiceResponse, Say
import time



# Your Account SID from twilio.com/console
account_sid = "coloque aqui"
# Your Auth Token from twilio.com/console
auth_token  = "coloque aqui "
client = Client(account_sid, auth_token)

# Abrir os 6 arquivos em Excel
lista_meses = ['quinto','sexto','setimo','oitavo','nono']
      
for mes in lista_meses:
    tabela_vendas = pd.read_excel(f'{mes}.xlsx')
    if(tabela_vendas['Notas'] >= 118).any():
      
      alunos = tabela_vendas.loc[tabela_vendas['Notas'] > 100, 'Category'].values[0]
      vendas = tabela_vendas.loc[tabela_vendas['Notas'] > 100, 'Notas'].values[0]
      print(f'Nesse bimestre os alunos nota 11 do {mes}  que estão como vencedores são : {alunos}, Nota: {vendas}')
      message = client.messages.create(
              to="+5588988848251",
              from_="+15594130353",
              body=f'Nesse bimestre os alunos nota 11 do {mes}  que estão como vencedores são : {alunos}, Nota: {vendas}')
      print(message.sid)

      response = VoiceResponse()
      response.say(f"""Os alunos  {mes} que alcançou título de aluno nota 11, Eu já enviei também um SMS com o nome do Vencedor, foi: {alunos}, Notas: {vendas}.""", voice='woman', language='pt-BR')
      response.say(f"""Estimados alunos nota 11:{alunos},você merece uma homenagem, espere um instante,""", voice='woman', language='pt-BR')
      response.play('https://programacao-fullstack.herokuapp.com/armando.mp3')
  
        

#mensagem = f"""

#<Response>
#<Say language='br-BR' utf8='pt-BR'>
#No mês {mes} alguém bateu a meta. Vendedor: {vendedor}, Notas: {vendas}
#</Say>
#<Play>https://ceece.farce.com.br/armando.mp3</Play>
#</Response>

#"""
      ligacao = client.calls.create(
          to="coloque aqui",
          from_="coloque aqui",
        twiml=response
      )
   
  
    