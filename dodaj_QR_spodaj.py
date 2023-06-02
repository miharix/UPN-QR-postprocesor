#Spacal: Miharix
#postprocesor za PDF generiran opomin da dolepi za bolj jasno kaj napisat na UPN in plačilno QR
#verzija 1.0
#pip install PyPDF2 reportlab reportlab_qrcode xlrd openpyxl --user
#pip install PyPDF2
#pip install reportlab
#pip install reportlab_qrcode
#pip install xlrd
#pip install openpyxl

import os
import glob
from PyPDF2 import PdfWriter, PdfReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen.canvas import Canvas
from reportlab_qrcode import QRCodeImage
from reportlab.lib.units import mm
import qrcode
import locale

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from datetime import datetime

import pandas as XLSX_opominov

from reportlab.lib.pagesizes import portrait,A4

Napake = 0

# IBAN podatki, kamor se naj plačilo nakaže
### podatki so le primer, to ni moj delodajalec
IBAN_Prejemnika = 'SI56 0400 1004 6484 329'

Ime_Prejemnika = 'Zveza za tehnično kulturo Slovenije'
Ulica_Prejemnika = 'Zaloška 65'
Posta_prejemnika = '1000 Ljubljana'

prebrano = XLSX_opominov.read_excel('Obdelaj/opomini.xlsx',header=0)
#print(prebrano)

locale.setlocale(locale.LC_ALL, 'sl_SI.UTF-8')
pdfmetrics.registerFont(TTFont('osifont', 'osifont.ttf'))

izvorni_PDF = PdfReader(open('Obdelaj/opomini.pdf', "rb"))
izhodni_PDF = PdfWriter()

#za vsak vnos v excelu naredi
for OPOMIN in range(len(prebrano)):

    Znesek = str(prebrano['Znesek'][OPOMIN])#'0' #[ZNESEK]
    Znesek = float(int(float(Znesek)*100)/100) #poskrbi da bosta samo 2 decimalki in float
    
    Namen_Placila = prebrano['Namen_placila'][OPOMIN]#Poravnava stroškov [Št. odločbe]
    Rok_Placila = prebrano['Rok_placila'][OPOMIN] #'2023-05-19' #[DAT_VALUTE]
    Koda_Namena = prebrano['Koda_namena'][OPOMIN]

    Referenca_Prejemnika = 'SI 00 '+ str(prebrano['Referenca_prejemnika'][OPOMIN])[:22] 
    Referenca_Prejemnika_QR = 'SI00'+ str(prebrano['Referenca_prejemnika'][OPOMIN])[:22]

    if (len(Referenca_Prejemnika_QR) > 26):
        print("!!!!!!! Referenca prejemnika je daljša kot 26 znakov")
        Napake += 1

    if (len(Namen_Placila) > 42):
        print("!!!!!!! Namen plačila je daljša kot 42 znakov")
        Napake += 1
        
    Ime_Placnika = prebrano['Ime_Priimek'][OPOMIN]
    Ulica_Placnika = prebrano['Naslov'][OPOMIN]
    Posta_Placnika = str(prebrano['Pošta'][OPOMIN])+' '+prebrano['Mesto'][OPOMIN]

    Rok_Placila = datetime.strptime(Rok_Placila.strftime('%d.%m.%Y'),'%d.%m.%Y' ).strftime('%d.%m.%Y')

    print("Pripravljam " + str(OPOMIN+1) + ". QR za: "+Ime_Placnika)
    #print(Znesek)
   # print(int(float(Znesek)*100))
   # print(locale.currency(float(int(float(Znesek)*100)/100), grouping=True))

    #v QR je znesek brez decimalk in z vodilnimi ničlami
    Znesek_za_QR = str(int(Znesek*100)).zfill(11)

    #pripravi besedilo in ga daj na polago

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=portrait(A4))

    t = can.beginText()
    t.setFont('osifont', 8)
    t.setCharSpace(1)
    t.setTextOrigin(350, 125)
    t.textLines(Ime_Placnika+",\nskenirajte to QR kodo\n za takojšno plačilo opomina v mobilni banki\n\nZnesek: "+ locale.currency(Znesek, grouping=True) +"\nKoda namena: " + Koda_Namena + "\nNamen plačila: " + Namen_Placila + "\nRok plačila: " + Rok_Placila + "\nIBAN prejemnika: "
                + IBAN_Prejemnika + "\nReferenca prejemnika: " + Referenca_Prejemnika + "\nNaslov prejemnika: " + Ime_Prejemnika + " " + Ulica_Prejemnika + " " + Posta_prejemnika + " ")
        
    can.drawText(t)

    ChecksumV_QR = len('UPNQR'+'\n')+len('\n\n\n\n')+len(Ime_Placnika + '\n')+len(Ulica_Placnika + '\n') + len(Posta_Placnika + '\n') + len(Znesek_za_QR + '\n') + len('\n\n') + len(Koda_Namena + '\n') + len(Namen_Placila + '\n') + len(Rok_Placila + '\n') + len(IBAN_Prejemnika.replace(' ', '') + '\n') + len(Referenca_Prejemnika_QR.replace(' ', '') + '\n') + len(Ime_Prejemnika + '\n') + len( Ulica_Prejemnika + '\n') + len(Posta_prejemnika + '\n')

    #dodaj QR
    vsebina_QR = 'UPNQR\n\n\n\n\n'+Ime_Placnika[:33] + '\n' + Ulica_Placnika[:33] + '\n' + Posta_Placnika[:32] + '\n' + Znesek_za_QR + '\n\n\n' + Koda_Namena[:4] + '\n' + Namen_Placila[:42] + '\n' + Rok_Placila[:10] + '\n' + IBAN_Prejemnika.replace(' ', '') + '\n' + Referenca_Prejemnika_QR.replace(' ', '') + '\n' + Ime_Prejemnika + '\n' + Ulica_Prejemnika + '\n' + Posta_prejemnika + '\n'  + str(ChecksumV_QR) + '\n'

    qr = QRCodeImage(vsebina_QR, size=35 * mm,fill_color='black',back_color=None,border=2 * mm ,version=15,error_correction=qrcode.constants.ERROR_CORRECT_M)
    qr.drawOn(can, 250, 22)

    can.save()

    #move to the beginning of the StringIO buffer
    packet.seek(0)

    # create a new PDF with Reportlab
    NoviPDF_buffer = PdfReader(packet)
    # Naloži obstoječi pdf
    #print(izvorni_PDF)
   
   
    # add the "watermark" (which is the new pdf) on the existing page
    try:
        papir = izvorni_PDF.pages[OPOMIN]
        papir.merge_page(NoviPDF_buffer.pages[0])
    except:
        print("\n!!!!!!! V opomini.xlsx je več vrstic kot je strani v opomini.pdf")
        Napake += 1
        continue
    izhodni_PDF.add_page(papir)

    
#kreiraj mapo če je še ni
os.makedirs("Dodan_QR", exist_ok=True)
 

print("\nShranjujem opomini-QR.pdf")

try:
    output_stream = open("Dodan_QR/opomini-QR.pdf", "wb")
    izhodni_PDF.write(output_stream)
    output_stream.close()
    print("Shranil: /Dodan_QR/opomini-QR.pdf\n")
except:
      print( "\n!!!!!!!!  NAPAKA: Datoteke /Dodan_QR/opomini-QR.pdf ni možno odpreti. A jo imate že odprto v drugem pogramu?\n")
      Napake += 1
      #continue



print("Končal vse kar sem našel")
if(Napake > 0):
    print("\n!!!! Število napak: "+str(Napake))
    input("\nPreglej napake in pritisni Enter za nadaljevanje...")
