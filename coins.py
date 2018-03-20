import urllib.request,json,os
import webbrowser,datetime
from hdata import hData


class coin:
	def __init__(self):
		self.baseurl = 'https://api.coinmarketcap.com/v1/ticker/'
		self.zeburl = 'https://www.zebapi.com/api/v1/market/ticker/btc/inr'
		# self.buyucoin = 'https://buyucoin.com/api/v1/btc'
		# self.unocoin = 'https://min-api.cryptocompare.com/data/price?fsym=BTC&tsyms=INR&e=Unocoin'
		self.date = datetime.date.today().strftime('%d-%m-%Y')
		self.time = datetime.datetime.now().time().strftime('%I:%M %p')

	def data(self):
		file = open('coin.htm','w')
		try:
			url = self.baseurl + '?' + urllib.parse.urlencode({'convert':'INR','limit':'15'})
			response = urllib.request.urlopen(url).read()
			data = json.loads(response)
			
			res_zebpay = urllib.request.urlopen(self.zeburl).read()
			data_zebpay = json.loads(res_zebpay)

			''' res_buyucoin = urllib.request.urlopen(self.buyucoin).read()
			data_buyucoin = json.loads(res_buyucoin)
			'''
			''' for coins in range(4):
				file.write('#Rank : '+ str(data[coins]['rank']) + '<br>')
				file.write('<b>'+str(data[coins]['name']) +'('+ str(data[coins]['symbol']) +')'+ '</b><br>')
				file.write('&nbsp&nbsp&nbsp&nbsp<font color="red">Price: ' + str(round(float(data[coins]['price_inr']),2)) +'</font>'+ '<br>')
				file.write('&nbsp&nbsp&nbsp&nbspPercent change in 24h: ' + str(data[coins]['percent_change_24h'])+'<br>')
				file.write('&nbsp&nbsp&nbsp&nbspPercent change in 7d: ' + str(data[coins]['percent_change_7d'])+'<br><br> '''
				


			file.write('<body style="background-color:#424949">**<font color=#5DADE2>'+self.date+'&nbsp||&nbsp'+self.time+'</font>**') 
				# <div>Exchange buy price: '+ data_zebpay['buy'] +' || Exchange sell price: '\
				#		+ data_zebpay['sell'] +'</div>')
			# file.write('**<font color=#A93226>' + self.date + '</font>**')
			# In tabular form
			file.write('<style>td{text-align:center;background-color:#52BE80} th{background-color:#45B39D} div{float:right;color:#5DADE2} .neg{color:#B03A2E;} .pos{color:#094710;} \
						.arw{margin-left:20px;} </style>' )# all CSS
			file.write('<br><table id="tb" border="1" align="center" style="width:100%"')
			file.write('<tr><th>Rank</th><th>Coin Name</th><th>Symbol</th><th>Price(INR)</th><th>Percent change in 1h</th><th>Percent change in 24h</th><th>Percent change in 7d</th></tr>')
			
			for coins in range(14):
				file.write('<tr><td>#' + str(data[coins]['rank']) + '</td>')
				file.write('<td>' + str(data[coins]['name']) + '</td>')
				file.write('<td>' + str(data[coins]['symbol']) + '</td>')
				file.write('<td><b><font color=#B03A2E>' + str(round(float(data[coins]['price_inr']),2)) +'</font></b></td>')
				
				if float(data[coins]['percent_change_1h']) < 0 :
					file.write('<td class="neg">' + str(data[coins]['percent_change_1h']) + '<span class="arw">&#9660 </span></td>')
				else:
					file.write('<td class="pos">' + str(data[coins]['percent_change_1h']) + '<span class="arw">&#128314 </span</td>')
				if float(data[coins]['percent_change_24h']) < 0 :
					file.write('<td class="neg">' + str(data[coins]['percent_change_24h']) + '<span class="arw">&#9660 </span></td>')
				else:
					file.write('<td class="pos">' + str(data[coins]['percent_change_24h']) + '<span class="arw">&#128314 </span></td>')
				if float(data[coins]['percent_change_7d']) < 0 :
					file.write('<td class="neg">' + str(data[coins]['percent_change_7d']) + '<span class="arw">&#9660 </span></td>')
				else:
					file.write('<td class="pos">' + str(data[coins]['percent_change_7d']) + '<span class="arw">&#128314 </span></td>')
				
			file.write('</table>') 
			if os.path.exists('histdata.csv'):
				file.write('<br><a href="histdata.csv" style="color:#3498DB">Download BTC historic data in CSV format<br>(2010 to Current date)</a>')
			else:
				self.historicData()
				file.write('<br><a href="histdata.csv" style="color:#3498DB">Download BTC historic data in CSV format<br>(2010 to Current date)</a>')

		except Exception as e:
			file.write('<br>Error!! '+ str(e))
		file.close()

	def historicData(self):
		hdata_obj = hData()
		hdata_obj.histdata()


if __name__ == '__main__':
	obj = coin()
	obj.data()
	webbrowser.open('coin.htm')