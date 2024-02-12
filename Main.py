import os.path
import math
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1gBlloegTsgqTe9OCz8wwPrNMtT4E9UPI4fxsb1OnvY8"
SAMPLE_RANGE_NAME = "engenharia_de_software!A2:F27"


def judge_situation(n1:int,n2:int,n3:int, faltas, total_aulas)-> str:
  average = take_average(n1,n2,n3)
  situation = "Aprovado"

  if faltimeter(total_aulas, faltas) > 25:
    situation = "Reprovado por falta"
    return situation

  if average < 50:
    situation = "Reprovado por nota"
  elif average < 70:
    situation = "Exame final"
  return situation

def take_average(n1:int,n2:int,n3:int):
  return float((n1 + n2 + n3)/3)

def missing_note(average):
  #5 <= (m + naf)/2
  #10 - m <= naf
  note = 100 - average
  note = math.ceil(note)
  return note

def faltimeter(classNumber, fouls)->int:
  percent = int((int(fouls)/int(classNumber))*100)
  return percent


def add(mat,situation, note):  
  creds = None 
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    # Adicionar/editar informação
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()
    row = int(mat)+3
    row = "G"+str(row)
    new_values = [
      [situation, note]
    ]
    result = (sheet.values()
              .update(
                spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                range=str(row), 
                valueInputOption="USER_ENTERED",#RAW or USER_ENTERED
                body={'values': new_values})
              .execute())
  except HttpError as err:
    print(err)

def main():
  # Autentication for integration with the googlesheets
  creds = None 
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("sheets", "v4", credentials=creds)

    # Reading data from the googleSheets
    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()        
    )

    NumberOfClasses = str(result['values'][0])[-4:-2]
    for item in result['values'][2:]:
      mat = item[0]
      nome = item[1]
      faltas = item[2]
      n1 = item[3]
      n2 = item[4]
      n3 = item[5]
      situation = judge_situation(int(n1),int(n2),int(n3), faltas, NumberOfClasses)
      if situation == "Exame final":
        notaFaltante = missing_note(take_average(int(n1),int(n2),int(n3)))
      else:
        notaFaltante = 0
      print(str(item) + "- Situação: "+ situation + "- Nota Faltante: " + str(notaFaltante))

      #updating with new informations
      add(mat,situation,notaFaltante)
  except HttpError as err:
    print(err)

if __name__ == "__main__":
  main()