# UPN QR postprocesor
 Doda na obstoječ večstranski PDF, QR kode skladne z UPN QR

Podatke za QR dobi iz excel datoteke

## Zakaj sem to spacal & Kdo to rabi ?
*Razvijalec računovodskega programa, je od nas zahteval absurdno visoko cifro za dogradnjo te funkcionalnosti.*
*Želja je bila, da se na račun/nalog/izvršbo/opomin/itd. doda čistopis podatkov za izpolnit položnico ter QR za digitalno plačilo.*

Torej če pošiljaš večim strankam na enkrat račune/opomine/...
iz katerih ni lepo pregledno vidno, koliko, komu nakazat, je ta mini posprocestor skriptica ravno zate.

Na PDF ti bo na dno dodala čistopis kaj naj plačnik napiše na položnico ter 
QR kodo, preko katere je možno plačat z uporabo mobilne banke.

# ! UPORABA SKRIPTE IN NJENEGA REZULTATA NA LASTNO ODGOVORNOST !
## ! Dobro pretestiraj, preden začneš pošiljat strankam !

# Namestitev
Testiran OS: Win10

1. Namesti Python 3 https://www.python.org/downloads/

	Za delovanje skripte je dovolj namestitev Pythona v user space (ne kot admin)

2. Restartaj PC

3. V CMD zaženi ukaz

	```
	pip install PyPDF2 reportlab reportlab_qrcode xlrd openpyxl --user`
	```

	ali pa enega po enega

	```
	pip install PyPDF2
	pip install reportlab
	pip install reportlab_qrcode
	pip install xlrd
	pip install openpyxl
	```

4. Zraven skripte v isto mapo skopiraj pisavo [osifont.ttf]
(vir: https://github.com/hikikomori82/osifont)

5. odpri dodaj_QR_spodaj.py
	* popravi IBAN, nasziv, naslov, ulico, 

6. Kreiraj 2 mapi
	* Dodan_QR
	* Obdelaj

# Uporaba

1. V mapo "Obdelaj" shrani
	* PDF datoteko na katerere strani želiš imeti QR kode imenovano opomini.pdf
	
		! Vsak račun/opomin/... naj ne bo daljši kot 1 stran

	* xlsx datoteko s seznamom polj za UPN, pazi imena naj bodo točno taka in v takem zaporedju

	```
	Ime_Priimek	Naslov	Pošta	Mesto	Znesek	Koda_namena	Namen_placila	Rok_placila	Referenca_prejemnika
	```

2. Dvoklikni dodaj_QR_spodaj.py

3. Počakaj da konzolno okno zgine.

4. V mapi Dodan_QR se pojavi verzija z dodano QR kodo.

5. ČE konzolno okno NE zgine, dobro poglej kako napako javlja 

# PAZI program ne preverja vrstnega reda
Stran 1 opomina -> dobi QR z podatki 1 vrstice excela itd.. 