#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Created on 18/05/2022

@author: JMichel
"""

''' https://www.google.com/settings/security/lesssecureapps '''

''' bibliothèques '''
from openpyxl import load_workbook
from datetime import datetime
import codecs, time
from mailjet_rest import Client
#

''' variables '''
#
''' date du jour '''
now = datetime.now()
today = now.strftime("%d/%m/%Y")
#
''' adresse mail & mot de passe expediteur '''
senderMail = 'sfxo6998@assoaris.org'
senderPwd = 'Aris-1982*'
#
''' serveur mail & port d'envoi '''
smtpServer = 'ares.o2switch.net'
smtpPort = 465
#
mailSubject = "ARIS - votre Anniversaire"
mailMessage = ""
textMessage = ""
htmlMessage = ""
#
''' fichiers '''
fchXLS = "./ARIS-MailingList-20221202.xlsx"
feuilXLS = "Adherents"
feuilLOG = "log"
feuilERR = "err"
fchTXT = "./ARIS-AnniversairesMails.txt"
fchHTML = ""
fchPDF = ""
#

''' fonctions '''
#
def clear():
    print("\033[H\033[J")
#
def setGlobal():
    global p,q,chk
    p , q = 0 , 0
    ''' variable de controle d'execution '''
    global chk
    chk = True
#
def showProcess(stp):
    global p
    p = p + 1
    print("-step:" +str(p) +" - " +stp +" -ok-")
#
def showDispatch(adr):
    global q
    q = q + 1
    print("-mail:" +str(q) +" - " +adr +" -ok-")
#

def writeText(txt):
    from os.path import isfile
    try:
        if(isfile(fchTXT)):
            fchW = codecs.open(fchTXT,"a",encoding="utf-8")
            fchW.write(txt)
            fchW.close()
    except:
        print("! erreur ecriture fichier " +fchTXT +" !")
        global chk
        chk = False
#

def getHtmlContent(prenom,nom):
   html = "<!DOCTYPE html>"
   html += "<html lang='fr'>"
   html += "<head>"
   html += "<meta charset='UTF-8'>"
   html += "<meta http-equiv='X-UA-Compatible' content='IE=edge'>"
   html += "<meta name='viewport' content='width=device-width, initial-scale=1.0'>"
   html += "<meta name='author' content='ARIS -Association des Anciens de la Banque Indosuez devenue CA-CIB- 12 place des Etats-Unis 92120 Montrouge'>"
   html += "<meta name='description' content='private notification, private information'>"
   html += "<title>ARIS Info</title>"
   html += "</head>"
   html += "<body>"
   html += "<div>cher(e) &nbsp;" + prenom
   html += "<br/><br/>Les membres du Conseil d'Administration de l'ARIS vous souhaitent un joyeux Anniversaire."
   html += "<br/><br/>bien amicalement."
   html += "</div>"
   html += "<br/><br/>"
   html += "<div style='font-size:x-small;text-align:center'>"
   html += "- mail non-commercial émanant de l'association <a href='http://www.anciensindosuez.org'>ARIS</a> à destination de ses adhérents -"
   html += "<div/>"
   html += "<div style='font-size:xx-small;text-align:center'>"
   html += "<span>ARIS -Association des Anciens de la Banque Indosuez devenue CA-CIB- , 12 place des Etats-Unis , 92120 Montrouge</span>"
   html += "<br/>"
   html += "<a href='mailto:anciensindosuez.aris@gmail.com?subject=Désinscription Mailing&body=Veuillez me désinscrire à votre mailing. Merci. "+ prenom +" " + nom + "'>désinscription mailing</a>"
   html += "<div/>"
   html += "</body>"
   html += "</html>"
   #
   return(html)
#

def openExcelFile():
    ''' ouverture du fichier Excel et selection du feuillet de données '''
    try:
        global workBook
        workBook = load_workbook(fchXLS)
        global workSheet
        workSheet = workBook[feuilXLS]
        showProcess("workBook & workSheet")
    except:
        print("! erreur de connexion au fichier Excel !")
        chk = False
        print("-chk:",chk)
#

def closeExcelFile():
    ''' fermeture du fichier Excel '''
    print("\n")
    try:
        workBook.close()
        showProcess("workBook closing")
    except:
        print("! erreur de fermeture du fichier Excel !")
#

''' MailJet '''
#
def execMailing():
   ''' lancement du processus d'execution d'envoi des mails à partir du fichier Excel '''
   print("\n")
   try:
        i = 2
        global anniversairesJour
        anniversairesJour = "\nanniversaires du " +today +" :"
        global workSheet, receiverMail
        while(workSheet["C"+str(i)].value != None):
           #print("i:", i)
           if(workSheet["C"+str(i)].value != "??"
           and (workSheet["D"+str(i)].value == "R"
                or  workSheet["D"+str(i)].value == "CR")):
                ''' recuperation de l' adresse mail du destinataire '''
                receiverMail = workSheet["C"+str(i)].value
                ''' recuperation du nom et du prenom du destinataire '''
                nom = workSheet["A"+str(i)].value
                if(workSheet["B"+str(i)].value != None):
                    prenom = workSheet["B"+str(i)].value
                else:
                   prenom = ""
                #print(prenom +" " +nom +" : ")
                #
                ''' recuperation de la date de naissance '''
                naissance = workSheet["E"+str(i)].value
                if(naissance[:5] == today[:5]):
                    anniversairesJour += "\n" +naissance +"  " + prenom +" " +nom
                    anniversairesJour += "  " +receiverMail +"  " +"OK"
                    ''' add mailContent to email '''
                    mailHtml = getHtmlContent(prenom,nom)
                    ''' envoi du mail '''
                    try:
                       execMailJet(mailHtml,prenom)
                       showDispatch(receiverMail)
                       ''' temporisation entre envois '''
                       time.sleep(3)
                    except:
                        print("! erreur envoi mail : " + receiverMail + " !")
                #
           i = i + 1
        print("\nAnniversaires du Jour :", anniversairesJour)
        writeText(anniversairesJour)
   except:
        err = str(i) + " - " + nom + " " + prenom + " : " + receiverMail + " -- "
        print("! erreur processus envoi mailing : " + err + " !")
#
def execMailJet(mailHtml,prenom):
    api_key = '136b489d805bef2748e9712c9ccfdd45'
    api_secret = 'abf698772be933996b1fec2ecf28fbdc'
    mailJet = Client(auth = (api_key, api_secret), version='v3.1')
    mailData = { "Messages":[ { "From": {"Email":senderMail, "Name":"ARIS"}
                              , "To": [ {"Email":receiverMail, "Name":prenom} ]
                              , "Subject":mailSubject
                              , "TextPart":senderMail+" -->  "+receiverMail
                              ,  "HTMLPart": mailHtml
                              }
                            ]
                }
    result = mailJet.send.create(data = mailData)
    print(result.status_code)
    #print(result.json())


''' execution '''
#
def main():
    clear()
    setGlobal()
    print("- start mailing -\n")
    if(chk == True):
        openExcelFile()
    if(chk == True):
        execMailing()
    closeExcelFile()
    # stopMailServer()
    print("\n- end mailing -")
#
if(__name__ == "__main__"):
  main()
#
def exitScript():
    raise Exception("exit()")


