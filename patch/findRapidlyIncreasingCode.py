import sendMessageToSlack
import win32com.client
from datetime import datetime
import time

# Check Connection Status
def connectionCheck():
	instCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
	connectionFlag = instCpCybos.IsConnect
	if connectionFlag != 1 :
		print("Connection Flag : %s, Connection Fail" %  (connectionFlag))
		exit(1)
	else :
		print("Connection Flag : %s, Connection Success" % (connectionFlag))
	return (connectionFlag)

# Get Company Stock Data ("A034220")
def getStockData(stockCode, chartCode):
	instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
	instStockChart.SetInputValue(0, stockCode)							# Set what to get
	instStockChart.SetInputValue(1, ord('2'))							# Set how to decide data period
	instStockChart.SetInputValue(4, 60)									# Set how many
	instStockChart.SetInputValue(5, (2))								# Set dat type
	instStockChart.SetInputValue(6, ord(chartCode))						# Set data duration
	# Send Request
	instStockChart.BlockRequest()
	return (instStockChart)

# Print Result
def printStockData(instStockChart):
	numberOfData = instStockChart.GetHeaderValue(3)
	numberOfDataType = instStockChart.GetHeaderValue(1)
	for i in range(numberOfData) :
		for j in range(numberOfDataType) :
			print(instStockChart.GetDataValue(j, i), end = " ")
		print("")

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

		# 1.5초 정도 재워서, 대신 증권 서버에 부하가 되지 않고, 1시간 정도 마다 분석할 수 있도록 함
		time.sleep(1.5)

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