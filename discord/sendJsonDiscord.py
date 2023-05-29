import discord
import json
import pandas as pd
from dotenv import load_dotenv
import os

main_dir = os.path.dirname(os.path.abspath(__file__))  # Diretório do script main.py
resulted_json = os.path.join(main_dir, "../resultados.json")
config_env = os.path.join(main_dir, "config.env")

load_dotenv(config_env, override=True)
id_bot = os.getenv('ID_BOT')
id_canal = int(os.getenv('ID_CHANNEL'))

intents = discord.Intents.default()
intents.typing = False
intents.presences = False

# Ler os dados do arquivo JSON
with open(resulted_json, 'r') as json_file:
    data = json.load(json_file)

# Extrair os dados do JSON
name = data['name']
result = data['result']
dates = data['dates']

# Criar uma mensagem com os dados
message = "===================================" + "\n"
message += f"{name}:\n"
message += "Datas distintas: " + dates + "\n"
message += "Resultados:\n"

for item in result:
    date = item['DATE']
    dimension = item['Dimensão']
    diff = item['Diferença (%)']
    message += f"{date}: {dimension} - {diff}%\n"

if 'Ok' in name:
    message = name

# Inicializar o bot Discord
client = discord.Client(intents=intents)

# Evento de inicialização do bot
@client.event
async def on_ready():
    print(f'Bot conectado como {client.user.name}')

    # Enviar a mensagem no canal desejado
    channel = client.get_channel(id_canal)  # Substitua pelo ID do canal correto
    await channel.send(message)

    await client.close()

# Executar o bot Discord
client.run(id_bot)