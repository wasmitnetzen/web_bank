#!/usr/bin/python3
# -*- coding: UTF-8 -*-
'''
Holt Kreditkarten-Umsätze per Web-Scraping vom Webfrontend der DKB
(Deutsche Kreditbank), die für diese Daten kein HCBI anbietet.
Die Umsätze werden im Quicken Exchange Format (.qif) ausgegeben
oder gespeichert und können somit in eine Buchhaltungssoftware
importiert werden.

Geschrieben 2007 von Jens Herrmann <jens.herrmann@qoli.de> http://qoli.de
Benutzung des Programms ohne Gewähr. Speichern Sie nicht ihr Passwort
in der Kommandozeilen-History!

7.9.2008: Anstelle von xml.xpath, das von Ubuntu 8.04 nicht mehr unterstützt
wird, wird jetzt lxml.etree benutzt.

25.10.2008: Kleiner Fix wegen einer HTML-Änderung der DKB (table wird nicht mehr
gebraucht), dafür wird jetzt exportiertes CSV ausgewertet, das nicht auf eine
Seite Ausgabe beschränkt ist. Da kein HTML mehr geparsed wird, entfällt auch die
Abhängigkeit zu tidy und xpath.

27.12.2009: Auswahl der Kartennummer

23.11.2011: Fix des Session-Handlings durch eine Änderung der DKB

12.05.2012: Fix des Session-Handlings durch erneute Änderung der DKB (Akzeptieren von
Cookies ist erforderlich) Danke an Robert Steffens!

28.01.2018: Anpassung der kompletten Abfragen nach Interface-Änderungen durch die DKB

Benutzung: web_bank.py [OPTIONEN]

 -a, --account=ACCOUNT      Kontonummer des Hauptkontos. Angabe notwendig
 -c, --card=NUMBER          Die letzten 4 Stellen der Kartennummer, falls
                            mehrere Karten vorhanden sind.
 -p, --password=PASSWORD    Passwort (Benutzung nicht empfohlen,
                            geben Sie das Passwort ein, wenn Sie danach
                            gefragt werden)
 -f, --from=DD.MM.YYYY      Buchungen ab diesem Datum abfragen
                            Default: Vor 30 Tagen
 -t, --till=DD.MM.YYYY      Buchungen bis zu diesem Datum abfragen
                            Default: Heute
 -n, --nice                 Gebe die Daten in einer formatierten Tabelle zurück
 -o, --outfile=FILE         Dateiname für die Ausgabedatei
                            Default: Standardausgabe (Fenster)
 -v, --verbose              Gibt zusätzliche Debug-Informationen aus
'''

import sys, getopt
from datetime import datetime
from datetime import timedelta
from getpass import getpass
import urllib.request, urllib.error, urllib.parse, urllib.request, urllib.parse, urllib.error, http.cookiejar, re

def group(lst, n):
	return list(zip(*[lst[i::n] for i in range(n)]))

debug=False
def log(msg):
	if debug:
		print(msg)

def debugHtmlToFile(content):
	with open("current.html", 'w') as htmlFile:
		htmlFile.write(content)


# Parser for new style banking pages
class NewParser:
	URL = "https://www.dkb.de"
	BETRAG = 'frmBuchungsbetrag'
	ZWECK = 'frmVerwendungszweck'
	TAG = 'frmBuchungstag'
	PLUSMINUS = 'frmSollHabenKennzeichen'
	MINUS_CHAR='S'
	DATUM = 'frmBelegdatum'
	ORGBETRAG = 'frmOriginalBuchungsbetrag'

	def get_cc_index(self, card, data):
		log('Finde Kreditkartenindex für Karte ***%s...'%card)
		pattern = '"(.)"[ ]*>[0-9\*]*{} / Kreditkarte'.format(card)
		index= re.findall(pattern, data)
		if len(index)>0:
			log('Index ist %s'%index[0])
			return index[0]
		else:
			print('Karte {} nicht gefunden!'.format(card))
			return '1'

	def get_cc_csv(self, account, card, password, fromdate, till, transactionStatus = '0'):
		log('Hole sessionID und Token...')
		# retrieve sessionid and token
		url= self.URL+"/banking"
		cj = http.cookiejar.LWPCookieJar()
		if debug:
			opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj),urllib.request.HTTPSHandler(debuglevel=1))
		else:
			opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj),urllib.request.HTTPSHandler(debuglevel=0))
		opener.addheaders = [('User-agent', 'Mozilla/5.0')]
		urllib.request.install_opener(opener)
		page = urllib.request.urlopen(url,).read().decode('utf-8')
		if "wartung_content" in page:
			raise Exception("Wartungsarbeiten bei DKB")
		token = re.findall('<input type="hidden" name="token" value="(.*)" id=',page)
		if token is None or len(token) == 0:
			raise Exception('Token not found on page {}'.format(page))
		token = token[0]
		sID = re.findall('<input type="hidden" name="\$sID\$" value="(.*)" ',page)[0]
		log('Token: {}, sID {}'.format(token,sID))
		# login
		url= self.URL+'/banking'
		request=urllib.request.Request(url, data=urllib.parse.urlencode({
		                                                     'token': token,
		                                                     '$sID$': sID,
		                                                     'j_username': account,
		                                                     'j_password': password,
		                                                     'browserName': "Firefox",
		                                                     'browserVersion': "40",
		                                                     '$event': 'login'
		                                                     }).encode('utf-8'))
		page=urllib.request.urlopen(request).read()
		referer = url
		url= self.URL+'/banking/finanzstatus/kreditkartenumsaetze'

		# retrieve data
		slAllAccounts = "1"
		request=urllib.request.Request(url, data= urllib.parse.urlencode({
		                                                     'slAllAccounts': slAllAccounts,
		                                                     'slSearchPeriod': '0',
		                                                     'filterType': 'DATE_RANGE',
		                                                     'postingDate': fromdate,
		                                                     'toPostingDate': till,
		                                                     '$event': 'search',
		                                                     'slTransactionStatus': transactionStatus

		}).encode('utf-8'), headers={'Referer':urllib.parse.quote_plus(referer)})
		data= urllib.request.urlopen(request).read().decode('utf-8')
		if card:
			slAllAccounts = self.get_cc_index(card,data)
			request=urllib.request.Request(url, data= urllib.parse.urlencode({
		                                                     'slAllAccounts': slAllAccounts,
		                                                     'slSearchPeriod': '0',
		                                                     'filterType': 'DATE_RANGE',
		                                                     'postingDate': fromdate,
		                                                     'toPostingDate': till,
		                                                     '$event': 'search',
		                                                     'slTransactionStatus': transactionStatus

			}).encode('utf-8'), headers={'Referer':urllib.parse.quote_plus(referer)})
			urllib.request.urlopen(request)



		#fetch CSV
		request=urllib.request.Request(url+'?$event=csvExport', headers={'Referer':urllib.parse.quote_plus(url)})
		antwort= urllib.request.urlopen(request).read().decode('iso-8859-1')
		log('Daten empfangen. Länge: {}, Typ: {}'.format(len(antwort),type(antwort)))
		return antwort

	def parse_csv(self, cc_csv):
		result=[]
		for line in cc_csv.split('\n')[8:]: # Liste beginnt in Zeile 9 des CSV
			g= line.split(';')
			if len(g)==7: #Jede Zeile hat 7 Elemente
				act={}
				act[self.ZWECK]=g[3][1:-1]
				act[self.TAG]=g[1][1:-1]
				act[self.DATUM]=g[2][1:-1]
				act[self.PLUSMINUS]=''
				act[self.BETRAG]=g[4][1:-1]
				act[self.ORGBETRAG]=g[5][1:-1]
				result.append(act)
		return result

	def render_csv(self,csv):
		render = "Wertstellung | Belegdatum |                                       Beschreibung | Betrag (EUR) |   Org. Betrag |\n"
		for line in csv:
			render += "{:>12s} | {:>9s} | {:>50s} | {:>12s} |  {:>12s} |\n".format(
				line[self.TAG],line[self.DATUM],line[self.ZWECK],line[self.BETRAG],line[self.ORGBETRAG])
		return render

CC_NAME= 'VISA'
CC_NUMBER= ''
LOGIN_ACCOUNT=''
LOGIN_PASSWORD=''
PARSER= NewParser()

GUESSES=[
		(PARSER.BETRAG,'-150.0','Aktiva:Barvermögen:Bargeld'),
]

def guessCategories(f):
	for g in GUESSES:
		if g[1] in f[g[0]].upper():
			return g[2]

def render_qif(cc_data):
	cc_qif=[]
	cc_qif.append('!Account')
	cc_qif.append('N'+CC_NAME)
	cc_qif.append('^')
	cc_qif.append('!Type:Bank')
	log('Für Ausgabe vorbereiten:')
	for f in cc_data:
		log(str(f))
		if PARSER.TAG in list(f.keys()):
			f[PARSER.BETRAG]= float(f[PARSER.BETRAG].replace('.','').replace(',','.'))
			if PARSER.MINUS_CHAR in f[PARSER.PLUSMINUS]:
				f[PARSER.BETRAG]= -f[PARSER.BETRAG]
			f[PARSER.BETRAG]=str(f[PARSER.BETRAG])
			datum=f[PARSER.DATUM].split('.')
			cc_qif.append('D'+datum[1]+'/'+datum[0]+'/'+datum[2])
			cc_qif.append('T'+f[PARSER.BETRAG])
			if PARSER.ZWECK+"1" in f:
				for n in range(1,8):
					if f[PARSER.ZWECK+str(n)].strip():
						cc_qif.append('M'+f[PARSER.ZWECK+str(n)])
			else:
				cc_qif.append('M'+f[PARSER.ZWECK])
			c= guessCategories(f)
			if c:
				cc_qif.append('L'+c)
			cc_qif.append('^')
	return '\n'.join(cc_qif)

class Usage(Exception):
	def __init__(self, msg):
		self.msg = msg

def main(argv=None):
	account=LOGIN_ACCOUNT
	password= LOGIN_PASSWORD
	card_no= CC_NUMBER
	fromdate=''
	till=datetime.now().strftime('%d.%m.%Y')
	outfile= sys.stdout

	if argv is None:
		argv = sys.argv
	try:
		try:
			opts, args = getopt.getopt(argv[1:], "ha:c:p:f:t:o:v:n", ['help','account=','card=','password=','from=','till=','outfile=','verbose','nice'])
		except getopt.error as msg:
			raise Usage(msg)
		for o, a in opts:
			if o in ("-h", "--help"):
				print(__doc__)
				return 0
			if o in ('-a','--account'):
				account= a
			if o in ('-c','--card'):
				card_no= a
			if o in ('-p','--password'):
				password= a
			if o in ('-f','--from'):
				fromdate= a
			if o in ('-t','--till'):
				till= a
			if o in ('-n','--nice'):
				formatted= True
			else:
				formatted= False
			if o in ('-o','--outfile'):
				try:
					outfile=open(a,'w')
				except IOError as msg:
					raise Usage(msg)
			if o in ('-v','--verbose'):
				print('Mit Debug-Ausgaben')
				global debug
				debug=True
		if not account:
			raise Usage('Kontonummer muss angegeben sein.')
		if not fromdate:
			fromdate = (datetime.now() - timedelta(days=30)).strftime('%d.%m.%Y')
			log("Anfangsdatum nicht gesetzt, setze 30 Tage = {}".format(fromdate))
		if not password:
			try:
				password=getpass('Geben Sie das Passwort für das Konto '+account+' ein: ')
			except KeyboardInterrupt:
				raise Usage('Sie müssen ein Passwort eingeben!')

		cc_csv = PARSER.get_cc_csv(account, card_no, password, fromdate, till)
		cc_data = PARSER.parse_csv(cc_csv)

		if formatted:
			print(PARSER.render_csv(cc_data))
		else:
			print(render_qif(cc_data), file=outfile)

	except Usage as err:
		print(__doc__, file=sys.stderr)
		print(err.msg, file=sys.stderr)
		return 2

if __name__ == '__main__':
	sys.exit(main())

