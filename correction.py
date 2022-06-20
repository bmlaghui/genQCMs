from functools import partial
from tkinter import *
from tkinter import filedialog
import math
import os
import time
from operator import itemgetter
import pymongo
import xlsxwriter as xlsxwriter
from tabula import read_pdf
from win32com import client


def ocr_responses(fname):
    df = read_pdf(fname, pages=1)
    data = df[1].values.tolist()
    responses = []
    infos = df[0].keys().tolist()
    infos += df[0].values.tolist()
    responses.append(infos[1])
    responses.append(infos[2][1])
    responses.append(infos[3][1])

    for x in data:
        length = len(x)
        middle_index = length // 2
        first_half = x[:middle_index]
        second_half = x[middle_index:]
        responses.append(first_half)
        responses.append(second_half)
    return responses


def ocr_answers(fname):
    df = read_pdf(fname, pages=1)
    data = df[0].values.tolist()
    data += df[1].values.tolist()

    return data


def calculate_note(fresponse, fanswers):
    responses = ocr_responses(fresponse)
    answers = ocr_answers(fanswers)
    name = responses[0]
    classe = responses[1]
    epreuve = responses[2]
    del responses[0]
    del responses[0]
    del responses[0]

    responses = sorted(responses, key=itemgetter(0))


    answersCorrect = []
    for item in answers:
        answer2 =  [1 if x=='X' else x for x in item]
        answer3 = [0 if math.isnan(x)  else x for x in answer2]
        answersCorrect.append(answer3[0:5])

    responsesCorrect = []
    for item in responses:
        item2 =  [1 if x=='X' else x for x in item]
        item3 = [0 if math.isnan(x)  else x for x in item2]
        responsesCorrect.append(item3)

    points = 0
    for i in range(0, len(answersCorrect)):
        if (responsesCorrect[i]==answersCorrect[i]):
            points +=1

    result = [name, classe, epreuve,  points/2]
    return result

def connect_to_db():
    try:
        myClient = pymongo.MongoClient("mongodb://localhost:27017/")
        myDB = myClient["iticQCM"]
    except:
        myDB = "Erreur de connexion"

    return myDB


def insert_to_db(db, notes):
    myCol = db["notes"]
    myCol.insert_one(notes)

def listFiles(directory):
    files = os.listdir(directory)
    for file in files:
        print(file)
    return files
def getNotes(db):
    collection_name = db["notes"]
    item_details = collection_name.find()
    for item in item_details:
        # This does not give a very readable output
        print(item['Nom et prénom'])
    return db;

def corriger_qcm(fname, dname):


    workbook = xlsxwriter.Workbook('/notes.xlsx')


    today = time.strftime('%d_%m_%Y')
    worksheet = workbook.add_worksheet('notes du '+ today)
    bold = workbook.add_format({'bold': True})

    worksheet.write(0, 3, "Feuille de notes éditée le "+time.strftime('%d/%m/%Y à %H:%M:%S'), bold)
    cell_format = workbook.add_format()
    cell_format.set_pattern(1)  # This is optional when using a solid fill.
    cell_format.set_bg_color('black')
    cell_format.set_color('white')
    cell_format.bold
    myDB = connect_to_db()
    col = 0
    row = 3

    for oneFile in dname:
        print(oneFile)
        responses = ocr_responses(oneFile)
        answers = ocr_answers(fname)
        notes = calculate_note(oneFile, fname)
        notesDic = dict.fromkeys(notes)
        keys = ["Nom et prénom", "Classe", "Epreuve", "Note"]
        notesDic = dict(zip(keys, notesDic))
        insert_to_db(myDB, notesDic)

        worksheet.write('A3', 'Nom et prénom', bold)
        worksheet.write('B3', 'Classe', bold)
        worksheet.write('C3', 'Epreuve', bold)
        worksheet.write('D3', 'Note', bold)
        print(row, col, notesDic["Nom et prénom"])
        worksheet.write(row, col, notesDic["Nom et prénom"])
        worksheet.write(row, col+1, notesDic["Classe"])
        worksheet.write(row, col+2, notesDic["Epreuve"])
        worksheet.write(row,col+3, notesDic["Note"])
        row = row+1


    worksheet.write(row+1, col + 0, "Moyenne générale", cell_format)
    worksheet.write_formula(row + 1, 1, '=AVERAGE(D2:D'+str(row)+')', cell_format)

    worksheet.write(row + 3, col + 0, "Note +", cell_format)
    worksheet.write_formula(row + 3, 1, '=MAX(D2:D' + str(row) + ')', cell_format)

    worksheet.write(row + 5, col + 0, "Note -", cell_format)
    worksheet.write_formula(row + 5, 1, '=MIN(D2:D' + str(row) + ')', cell_format)
    print('=SUM(D2:D'+str(row)+')')
    worksheet.set_column(0, 0, 25)
    workbook.close()


    input_file = os.path.abspath('/notes.xlsx')
    # give your file name with valid path
    output_file = os.path.abspath('/notes.pdf')
    # give valid output file name and path
    app = client.DispatchEx("Excel.Application")
    app.Interactive = False
    app.Visible = False
    Workbook = app.Workbooks.Open(input_file)
    try:
        Workbook.ActiveSheet.ExportAsFixedFormat(0, output_file)
    except Exception as e:
        print("Failed to convert in PDF format.Please confirm environment meets all the requirements  and try again")
        print(str(e))
    finally:
        Workbook.Close()
