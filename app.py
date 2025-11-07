import streamlit as st
import pandas as pd
import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from num2words import num2words

# === CONFIG ===
SPREADSHEET_ID = "INSERISCI_ID_DEL_TUO_GOOGLE_SHEET"  # es. 1AbCDeF...
SHEET_NAME = "Rubrica"

# === AUTENTICAZIONE GOOGLE ===
creds = service_account.Credentials.from_service_account_file(
    "credentials.json",
    scopes=["https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/documents"]
)

sheets_service = build('sheets', 'v4', credentials=creds)
docs_service = build('docs', 'v1', credentials=creds)
drive_service = build('drive', 'v3', credentials=creds)

# === FUNZIONE PER LEGGERE I DATI DALLO SHEET ===
def carica_rubrica():
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A2:H"
    ).execute()
    values = result.get('values', [])
    if not values:
        return pd.DataFrame()
    df = pd.DataFrame(values, columns=[
        "nome", "indirizzo", "codice_fiscale", "email", "pec", "tipo_operazione", "evento", "note"
    ])
    return df

# === FUNZIONE PER CREARE IL DOCUMENTO ===
def crea_documento(riga, importo, data, evento, luogo, numero):
    importo_lettere = num2words(float(importo), lang='it')

    # Testi base
    if riga["tipo_operazione"].lower().strip() == "ricevuto":
        corpo = f"""
Dichiaro di aver ricevuto a titolo di contributo da {riga['nome']},
{riga['indirizzo']},
C.F.: {riga['codice_fiscale']} – Email: {riga['email']} – PEC: {riga['pec']}

la somma di € {importo} (diconsi {importo_lettere} euro),
come contributo spese per l’esibizione di musica live in occasione
dell’evento “{evento}”, tenutasi a {luogo} il giorno {data}.

Documento n. {numero}.
"""
        titolo = f"Ricevuta_{numero}_{riga['nome']}"
    else:
        corpo = f"""
Dichiaro di aver versato a titolo di contributo a {riga['nome']},
{riga['indirizzo']},
C.F.: {riga['codice_fiscale']} – Email: {riga['email']} – PEC: {riga['pec']}

la somma di € {importo} (diconsi {importo_lettere} euro),
quale compenso per l’esibizione di musica live in occasione
dell’evento “{evento}”, tenutasi a {luogo} il giorno {data}.

Documento n. {numero}.
"""
        titolo = f"Versamento_{numero}_{riga['nome']}"

    # Crea documento
    doc = docs_service.documents().create(body={"title": titolo}).execute()
    doc_id = doc.get('documentId')

    # Inserisci il testo
    docs_service.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": [{"insertText": {"location": {"index": 1}, "text": corpo}}]}
    ).execute()

    # Restituisce link
    return f"https://docs.google.com/document/d/{doc_id}/edit"

# === INTERFACCIA STREAMLIT ===
st.title("🧾 Generatore Documenti Associazione")

df = carica_rubrica()
if df.empty:
    st.error("⚠️ Rubrica vuota o non trovata.")
else:
    nome = st.selectbox("Cerca un nome:", df["nome"].tolist())

    riga = df[df["nome"] == nome].iloc[0]

    st.markdown("---")
    st.subheader("Dati per il documento")
    importo = st.text_input("Importo (€)")
    data = st.date_input("Data", datetime.date.today())
    evento = st.text_input("Evento (facoltativo)", "")
    luogo = st.text_input("Luogo")
    numero = st.text_input("Numero documento")

    if st.button("📄 Genera documento"):
        if not importo or not luogo or not numero:
            st.warning("Compila almeno importo, luogo e numero documento.")
        else:
            link = crea_documento(riga, importo, data, evento, luogo, numero)
            st.success(f"Documento creato con successo! 👉 [Apri documento]({link})")
