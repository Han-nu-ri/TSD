import win32com.client
import pymysql
from datetime import datetime
import time

class CpConclusionEvent:
	def set_params(self, connection, cpConclusion):
		self.connection = connection
		self.cpConclusion = cpConclusion

	def OnReceived(self):
		insertStockLog(self.connection, self.cpConclusion)

class CpConclusion:
	def __init__(self, connection):
		self.connection = connection
		self.obj = win32com.client.Dispatch("DsCbo1.CpConclusion")

	def Subscribe(self):
		handler = win32com.client.WithEvents(self.obj, CpConclusionEvent)
		handler.set_params(self.connection, self.obj)
		self.obj.Subscribe()

	def Unsubscribe(self):
		sefl.obj.Unsubscribe()

def connectionCheck():
	instCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
	connectionFlag = instCpCybos.IsConnect
	if connectionFlag != 1 :
		print("Connection Flag : %s, Connection Fail" %  (connectionFlag))
		exit(1)
	else :
		print("Connection Flag : %s, Connection Success" % (connectionFlag))
	return (connectionFlag)

def initAccount():
	objTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")
	initCheck = objTrade.TradeInit(0)
	if initCheck != 0 :
		print("Account Init Fail")
		exit(1)
	else :
		print("Account Init Success")
	return (objTrade)

def setOrderParameter(objTrade, orderType, stockCode, volume, value):
	# Setting Order Parameter
	acc = objTrade.AccountNumber[0]		# 계좌번호
	objStockOrder = win32com.clinet.Dispatch("CpTrade.CpTd0311")
	objStockOrder.SetInputValue(0, orderType)	# 1 : 매도, 2 : 매수
	objStockOrder.SetInputValue(1, acc)
	objStockOrder.SetInputValue(3, stockCode)
	objStockOrder.SetInputValue(4, volume)
	objStockOrder.SetInputValue(5, value)
	return (objStockOrder)

def buyStock(objTrade, stockCode, volume, value):
	objStockOrder = setOrderParameter(objTrade, "2", stockCode, volume, value)
	# Request Order
	objStockOrder.BlockRequest()

def sellStock(objTrade, stockCode, volume, value):
	objStockOrder = setOrderParameter(objTrade, "1", stockCode, volume, value)

	# Request Order
	objStockOrder.BlockRequest()

# MySql Connection
def getDbConnection():
	ipAddress = '127.0.0.1'
	userId = 'root'
	userPassword = 'root'
	dataBase = 'stock'
	connection = pymysql.connect(host = ipAddress, user = userId, password = userPassword, db = dataBase, charset = 'utf8')
	return (connection)

# Insert Stock Log Data
def insertStockLog(connection, cpConclusion):
	# Get Cursor
	cursor = connection.cursor()
	dateOfData = datetime.now()
	stockCode = cpConclusion.GetHeaderValue(9)
	price = cpConclusion.GetHeaderValue(4)
	volume = cpConclusion.GetHeaderValue(3)
	orderType = cpConclusion.GetHeaderValue(12)	# 1 : sell, 2 : buy
	sqlCommand = """INSERT INTO StockLog(StockCode, DateOfData, Price, Volume, OrderType)
			VALUES (%s, %s, %s, %s, %s)"""
	cursor.execute(sqlCommand, (stockCode, dateOfData, price, volume, orderType))

	# Commit
	connection.commit()
	cursor.close()

# Get All Company Info
def getAllStockCode():
	instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
	codeList = instCpCodeMgr.GetStockListByMarket(1)
	stockCodeList = {}
	for code in codeList :
		companyName = instCpCodeMgr.CodeToName(code)
		stockCodeList[code] = companyName
	return stockCodeList

def findRapidlyIncreasing() :
	# 1. check connection check
	connectionCheck()

	# 2. Get All Stock Code
	stockCodeList = getAllStockCode()

	# Account Init : Enter Account Password
	objTrade = initAccount()

	for stockCode in stockCodeList.keys() :
		companyName = stockCodeList[stockCode]
		print("[rapid increasing 분석]\n현재 시각 : %s,  분석 대상 : %s" % (str(datetime.now()), companyName))

		# 3. get startPrice
		instStockChart = getStockData(stockCode, 'D')
		startPrice = instStockChart.GetDataValue(0, 0);

		# 4. get currentPrice
		instStockChart = getStockData(stockCode, 'm')
		currentPrice = instStockChart.GetDataValue(0, 0);

		# 5. calculate net change
		increasedRatio = ((currentPrice - startPrice) / startPrice) * 100

		# 6. check rapidly or not
		if increasedRatio > 5 :
			# Send Slack Message if rapidly increaded
			message = "[rapidly increasing]\n %s, 상승 비율 : %f%% \n시가 : %d, 현재가 : %d" % (companyName, increasedRatio, startPrice, currentPrice)
			sendMessageToSlack.sendMessage(message)
			buyStock(objTrade, stockCode, 10, currentPrice)

		# 1.5초 정도 재워서, 대신 증권 서버에 부하가 되지 않고, 1시간 정도 마다 분석할 수 있도록 함
		time.sleep(1.5)

# getDbConnection
connection = getDbConnection()

# CpConclusion Subscribe
cpConclusion = CpConclusion(connection)
cpConclusion.Subscribe()

# did this until forever
while True :
	# get current time and decide wheather execute or not
	currentTime = datetime.now()
	hour = int(currentTime.strftime('%H'))
	minute = int(currentTime.strftime('%M'))

	if hour >= 9 and hour <= 15 :
		findRapidlyIncreasing()
	else :
		print("장이 종료 되어, 프로그램을 종료합니다.\n종료 시각 : %s " % currentTime)
		exit(1)