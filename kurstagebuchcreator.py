# imports
from docxtpl import DocxTemplate #docx-editor package
import ferien #import ferien-api for Germany
import datetime
import sys
import time


# settings
bundesland = "BW" # possible values (Bundesland hier angeben): BW; BY; BE; BB; HB; HH; HE; MV; NI; NW; RP; SL; SN; ST; SH; TH; Siehe ISO-3166-2:DE


def checkVacation(date):
    #convert to tz=none
    realdate = date.astimezone(tz=None)

    #check vacation
    global bundesland
    try:
        ferien_data = ferien.state_vacations(bundesland)
    except:
        print("Fehler beim Laden der Ferien. Bitte Internetverbindung prüfen und Programm neu starten.")
        sys.exit()

    ferien_gefunden = False
    for fobj in ferien_data:
        if fobj.start < realdate < fobj.end:
            ferien_gefunden = True
            break
        else:
            pass
        if ferien_gefunden == True:
            break

    return(ferien_gefunden)
    
def getDates(beginningDate, endDate):
    # variables
    allDates = []

    # convert to datetime-objects
    beginningDate = datetime.datetime.strptime(beginningDate, "%d.%m.%Y")
    endDate = datetime.datetime.strptime(endDate, "%d.%m.%Y")

    # loop through all possible dates
    actDate = beginningDate
    while (actDate != endDate):
        if checkVacation(actDate) == False:
            allDates.append(actDate.strftime("%d.%m.%Y"))
        actDate = actDate + datetime.timedelta(days=7)

    allDates.append(endDate.strftime("%d.%m.%Y"))
    return(allDates)

def buildKursplan(kursname, ortsgruppe, stp, kursperiode, dlrgid, dateBeginning, dateEnd):
    # open document
    doc = DocxTemplate("data/template.docx")

    # write data into context variable
    context = {"kurs" : kursname, "og": ortsgruppe, "stp" : stp, "kp" : kursperiode, "dID" : dlrgid, "col_labels": ["Geplanter Kursinhalt", "Gemachter Kursinhalt", "Ausbilder"], "all_dates": []}
    dates = getDates(dateBeginning, dateEnd)
    for date in dates:
        preedit = {"label":date, "cols":["","",""]}
        context["all_dates"].append(preedit)
    
    # render and save document to output.docx
    doc.render(context)
    doc.save("[" + dlrgid + "] Kurstagebuch " + kursname + " " + kursperiode + ".docx")

# Main
input_kursname = input("Bitte Kursnamen eingeben: ")
input_ortsgruppe = input("Bitte Ortsgruppe angeben: ")
input_stuetzpunkt = input("Bitte Stützpunkt angeben: ")
input_kursperiode = input("Bitte Kursperiode angeben (Jahr+_Herbst oder _Frühjahr): ")
input_beginning = input("Bitte Kursstart angeben (tt.mm.jjjj): ")
input_end = input("Bitte letzten Kurstermin eingeben (tt.mm.jjjj): ")
input_did = input("Bitte DLRG-ID eingeben (KP-<Kursname>-01 wenn nichts Anderes angegeben): ")
print("Kursplan wird generiert... Bitte warten!")
buildKursplan(input_kursname, input_ortsgruppe, input_stuetzpunkt, input_kursperiode, input_did, input_beginning, input_end)
print("Kursplan erfolgreich erstellt. Er befindet sich in dem Ordner, wo dieses Programm gespeichert wurde!")
time.sleep(10)