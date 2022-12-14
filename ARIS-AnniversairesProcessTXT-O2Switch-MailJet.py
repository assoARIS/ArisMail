#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Created on 18/05/2022

@author: JMichel
"""

''' https://www.google.com/settings/security/lesssecureapps '''

import time, codecs
from datetime import datetime
from json import dumps, loads
from mailjet_rest import Client

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
fchTXT_R = "./ARIS-AnniversairesListe.txt"
fchTXT_W = "./ARIS-AnniversairesMails.txt"
fchPDF = ""
#

''' fonctions '''
#
def clear():
    print("\033[H\033[J")
#
def getDaytime():
    now = datetime.now()
    daytime = now.strftime("%d/%m/%Y %H:%M:%S")
    return daytime
#
def setGlobal():
    global p , q
    p = 1 ; q = 0
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

def readText():
    from os.path import isfile
    try:
        if(isfile(fchTXT_R)):
            fchR = codecs.open(fchTXT_R,"r",encoding="utf-8")
            txt = fchR.read()
            fchR.close()
            return(txt)
        else:
            return("??")
    except:
        print("! erreur lecture fichier " +fchTXT_R +" !")
        chk = False
#
def writeText(txt):
    from os.path import isfile
    try:
        if(isfile(fchTXT_W)):
            fchW = codecs.open(fchTXT_W,"a",encoding="utf-8")
            fchW.write(txt)
            fchW.close()
    except:
        print("! erreur ecriture fichier " +fchTXT_W +" !")
        chk = False
#

def getHtmlContent(nom,prenom):
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

''' MailJet '''
#
def execMailing():
    ''' lancement du processus d'execution d'envoi des mails à partir du fichier texte '''
    nom = "??" ; prenom = "??"
    try:
        global anniversairesJour, receiverMail
        anniversairesJour = "\nanniversaires du " +getDaytime() +" :"
        receiverMail = "@"
        global jsonAnniversaires, dicoAnniversaires
        ''' lecture du fichier texte listant les adhérents au format json '''
        jsonAnniversaires = readText()
        #print("jsonAnniversaires:\n",jsonAnniversaires)
        ''' conversion en dictionnaire '''
        dicoAnniversaires = loads(jsonAnniversaires)
        #print("dicoAnniversaires:\n",dicoAnniversaires)
        #print("\n")
        print("anniversaires du " +today +" :")
        ''' recherche et extraction des adhérents ayant une date d'anniversaire correspondant au jour présent '''
        for itm in dicoAnniversaires.items():
            naissance = itm[1][4]
            if(naissance[:5] == today[:5]):
                print(itm)
                nom = itm[1][0]
                prenom = itm[1][1]
                receiverMail = itm[1][2]
                anniversairesJour += "\n" +naissance +"  " +nom+" " +prenom
                anniversairesJour += "  " +receiverMail +"  " +"OK"
                ''' add mailContent to email '''
                mailHtml = getHtmlContent(nom,prenom)
                ''' envoi du mail '''
                try:
                   execMailJet(mailHtml,prenom)
                   showDispatch(receiverMail)
                   ''' temporisation entre envois '''
                   time.sleep(3)
                except:
                    print("! erreur envoi mail : " + receiverMail + " !")#
        print(anniversairesJour)
        writeText(anniversairesJour)
    except:
        err = nom + " " + prenom + " : " + receiverMail + " -- "
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
        execMailing()
    print("\n- end mailing -")
#   
if(__name__ == "__main__"):
  main()
#
def exitScript():
    raise Exception("exit()")

