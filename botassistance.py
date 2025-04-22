#!/usr/bin/python
# Bot Telegram per assistenza tecnica informatica con registrazione dati su Excel
import telebot
import pandas as pd
import os
from datetime import datetime
from telebot import types

API_TOKEN = '7602533521:AAHXgTzgQ8IyovywOWfrq3XtmzcMZllV6Nc'
bot = telebot.TeleBot(API_TOKEN)
EXCEL_FILE = 'assistenza_tecnica.xlsx'

# Dizionario per memorizzare informazioni temporanee durante la conversazione
user_data = {}

# Verifica se il file Excel esiste, altrimenti lo crea
def check_excel_file():
    if not os.path.exists(EXCEL_FILE):
        # Crea un dataframe vuoto con le colonne necessarie
        df = pd.DataFrame(columns=[
            'ID Utente', 'Username', 'Nome', 'Cognome', 'Cellulare', 'Email',
            'Data', 'Ora', 'Tipo Problema', 'Descrizione', 'Stato'
        ])
        # Salva il file
        df.to_excel(EXCEL_FILE, index=False)
        print(f"File {EXCEL_FILE} creato con successo.")
    else:
        print(f"File {EXCEL_FILE} già esistente.")

# Registra una richiesta nel file Excel
def registra_richiesta(user_id, username, nome, cognome, cellulare, email, tipo_problema, descrizione):
    now = datetime.now()
    data = now.strftime("%d/%m/%Y")
    ora = now.strftime("%H:%M:%S")
    
    # Leggi il file Excel esistente
    df = pd.read_excel(EXCEL_FILE)
    
    # Aggiungi la nuova richiesta
    nuova_richiesta = {
        'ID Utente': user_id,
        'Username': username,
        'Nome': nome,
        'Cognome': cognome,
        'Cellulare': cellulare,
        'Email': email,
        'Data': data,
        'Ora': ora,
        'Tipo Problema': tipo_problema,
        'Descrizione': descrizione,
        'Stato': 'Aperto'
    }
    
    # Aggiungi al dataframe e salva
    df = pd.concat([df, pd.DataFrame([nuova_richiesta])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)
    print(f"Richiesta registrata per l'utente {nome} {cognome}")

# Verifica la presenza del file Excel all'avvio
check_excel_file()

# Handle '/start' and '/help'
@bot.message_handler(commands=['help', 'start'])
def send_welcome(message):
    bot.reply_to(message, """\
Inanzitutto ciao e benvenuto al nuovo servizio di Assistenza Tecnica Informatica, questo è solo un esempio di utilizzo per assistenza tecnica mi chiamo Virgilio è sarò la tua guida 
Puoi inviare una richiesta utilizzando il comando:
/nuova_richiesta

Tipi di problemi disponibili:
- Hardware
- Software
- Rete
- Stampanti
- Account
- Altro

Per verificare lo stato delle tue richieste, usa:
/stato
""")

# Handle '/nuova_richiesta' command
@bot.message_handler(commands=['nuova_richiesta'])
def start_richiesta(message):
    chat_id = message.chat.id
    user_id = message.from_user.id
    
    # Inizializza i dati utente
    user_data[user_id] = {
        'username': message.from_user.username
    }
    
    # Chiedi il nome
    msg = bot.send_message(chat_id, "Per iniziare una nuova richiesta di assistenza, inserisci il tuo nome:")
    bot.register_next_step_handler(msg, process_nome_step)

def process_nome_step(message):
    try:
        user_id = message.from_user.id
        nome = message.text
        user_data[user_id]['nome'] = nome
        
        # Chiedi il cognome
        msg = bot.send_message(message.chat.id, f"Grazie {nome}, ora inserisci il tuo cognome:")
        bot.register_next_step_handler(msg, process_cognome_step)
    except Exception as e:
        bot.reply_to(message, f"Si è verificato un errore: {str(e)}")

def process_cognome_step(message):
    try:
        user_id = message.from_user.id
        cognome = message.text
        user_data[user_id]['cognome'] = cognome
        
        # Chiedi il numero di cellulare
        msg = bot.send_message(message.chat.id, "Inserisci il tuo numero di cellulare per essere contattato:")
        bot.register_next_step_handler(msg, process_cellulare_step)
    except Exception as e:
        bot.reply_to(message, f"Si è verificato un errore: {str(e)}")

def process_cellulare_step(message):
    try:
        user_id = message.from_user.id
        cellulare = message.text
        user_data[user_id]['cellulare'] = cellulare
        
        # Chiedi l'email
        msg = bot.send_message(message.chat.id, "Inserisci la tua email:")
        bot.register_next_step_handler(msg, process_email_step)
    except Exception as e:
        bot.reply_to(message, f"Si è verificato un errore: {str(e)}")

def process_email_step(message):
    try:
        user_id = message.from_user.id
        email = message.text
        user_data[user_id]['email'] = email
        
        # Crea tastiera inline per i tipi di problema
        markup = types.InlineKeyboardMarkup(row_width=2)
        hardware = types.InlineKeyboardButton("Hardware", callback_data="prob_Hardware")
        software = types.InlineKeyboardButton("Software", callback_data="prob_Software")
        rete = types.InlineKeyboardButton("Rete", callback_data="prob_Rete")
        stampanti = types.InlineKeyboardButton("Stampanti", callback_data="prob_Stampanti")
        account = types.InlineKeyboardButton("Account", callback_data="prob_Account")
        altro = types.InlineKeyboardButton("Altro", callback_data="prob_Altro")
        
        markup.add(hardware, software, rete, stampanti, account, altro)
        
        bot.send_message(message.chat.id, "Seleziona il tipo di problema:", reply_markup=markup)
    except Exception as e:
        bot.reply_to(message, f"Si è verificato un errore: {str(e)}")

@bot.callback_query_handler(func=lambda call: call.data.startswith('prob_'))
def callback_query(call):
    try:
        user_id = call.from_user.id
        tipo_problema = call.data.split('_')[1]
        user_data[user_id]['tipo_problema'] = tipo_problema
        
        bot.answer_callback_query(call.id)
        bot.send_message(call.message.chat.id, f"Hai selezionato: {tipo_problema}")
        
        # Chiedi la descrizione del problema
        msg = bot.send_message(call.message.chat.id, "Descrivi dettagliatamente il problema che stai riscontrando:")
        bot.register_next_step_handler(msg, process_descrizione_step)
    except Exception as e:
        bot.send_message(call.message.chat.id, f"Si è verificato un errore: {str(e)}")

def process_descrizione_step(message):
    try:
        user_id = message.from_user.id
        descrizione = message.text
        
        # Registra la richiesta
        registra_richiesta(
            user_id,
            user_data[user_id]['username'],
            user_data[user_id]['nome'],
            user_data[user_id]['cognome'],
            user_data[user_id]['cellulare'],
            user_data[user_id]['email'],
            user_data[user_id]['tipo_problema'],
            descrizione
        )
        
        # Risposta all'utente
        bot.send_message(message.chat.id, 
            f"Richiesta di assistenza registrata con successo!\n\n"
            f"Nome: {user_data[user_id]['nome']} {user_data[user_id]['cognome']}\n"
            f"Cellulare: {user_data[user_id]['cellulare']}\n"
            f"Email: {user_data[user_id]['email']}\n"
            f"Tipo problema: {user_data[user_id]['tipo_problema']}\n"
            f"Descrizione: {descrizione}\n\n"
            f"Un tecnico ti contatterà al più presto.")
        
        # Pulisci i dati temporanei
        del user_data[user_id]
    except Exception as e:
        bot.reply_to(message, f"Si è verificato un errore: {str(e)}")

# Handle '/stato' command to check request status
@bot.message_handler(commands=['stato'])
def check_status(message):
    try:
        # Leggi il file Excel
        df = pd.read_excel(EXCEL_FILE)
        
        # Filtra per l'utente corrente
        user_requests = df[df['ID Utente'] == message.from_user.id]
        
        if user_requests.empty:
            bot.reply_to(message, "Non hai nessuna richiesta di assistenza attiva.")
            return
            
        # Prepara la risposta
        risposta = "Le tue richieste di assistenza:\n\n"
        for index, row in user_requests.iterrows():
            risposta += f"Data: {row['Data']} - {row['Ora']}\n"
            risposta += f"Nome: {row['Nome']} {row['Cognome']}\n"
            risposta += f"Contatti: {row['Cellulare']} | {row['Email']}\n"
            risposta += f"Problema: {row['Tipo Problema']}\n"
            risposta += f"Descrizione: {row['Descrizione']}\n"
            risposta += f"Stato: {row['Stato']}\n"
            risposta += "------------------------\n"
            
        bot.reply_to(message, risposta)
        
    except Exception as e:
        bot.reply_to(message, f"Si è verificato un errore: {str(e)}")
        print(f"Errore: {str(e)}")

# Handle all other messages
@bot.message_handler(func=lambda message: True)
def echo_message(message):
    bot.reply_to(message, "Non ho capito il comando. Usa /help per vedere i comandi disponibili o /nuova_richiesta per aprire una nuova richiesta di assistenza.")

# Start the bot
print("Bot avviato!")
bot.infinity_polling()