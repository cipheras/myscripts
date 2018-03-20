'''
File writes original bitcoin price and zebpay sell and buy price in text file which is saved on desktop.
File also shows the difference in zepbay price and original price. 
'''
import urllib,json,urllib.request,datetime,os

class price:
	def __init__(self):
		self.zeburl = 'https://api.zebpay.com/api/v1/ticker?currencyCode=INR'
		self.btcurl = 'https://blockchain.info/ticker'
		self.date = datetime.date.today().strftime('%d-%m-%y')
	
	def log_struc(self):
		file = open('zebpay.txt','w')
		try:
			file.write("{: <18} {: <18} {: <18} {: <18} {: <18} {: <18}".format('BTC_price','Zebpay_Buy_price','Zebpay_Sell_price','BTC-Buy(L in -)','BTC-Sell(L in +)','Buy-Sell')+'\n')
			file.write('-------------------------------------------------------------------------------------------------------\n\n')
			
		except Exception as e:
			file.write('Error!! '+ str(e))
		file.close()

	def log(self):
		file = open('zebpay.txt','a')
		try:
			req_zeburl = urllib.request.urlopen(self.zeburl).read()
			req_btcurl = urllib.request.urlopen(self.btcurl).read()

			jsonData_zeb = json.loads(req_zeburl)
			jsonData_btc = json.loads(req_btcurl)

			d1 = jsonData_btc['INR']['last'] - jsonData_zeb['buy']
			d2 = jsonData_btc['INR']['last'] - jsonData_zeb['sell']
			d3 = jsonData_zeb['buy'] - jsonData_zeb['sell']

			file.write("{: <18} {: <18} {: <18} {: <18} {: <18} {: <18}".format(str(jsonData_btc['INR']['last']),str(jsonData_zeb['buy']),str(jsonData_zeb['sell']),str(round(d1,2)),str(round(d2,2)),str(round(d3,2)))+'\n')
			file.write('############################################### '+self.date+' ##############################################\n\n')
			print("Script writing log....")
		
		except Exception as e:
			file.write('\nError!! '+ str(e))
		file.close()

obj = price()
if  os.path.exists('zebpay.txt'):
	obj.log()
else :
	obj.log_struc()
	obj.log()