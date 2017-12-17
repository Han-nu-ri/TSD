import win32com.client
import pymysql

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
def getStockData(stockCode):
	instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
	instStockChart.SetInputValue(0, stockCode)							# Set what to get
	instStockChart.SetInputValue(1, ord('2'))							# Set how to decide data period
	instStockChart.SetInputValue(4, 1)									# Set how many
	instStockChart.SetInputValue(5, (0, 2, 3, 4, 5, 6, 7, 8, 9))		# Set dat type
	instStockChart.SetInputValue(6, ord('D'))							# Set data duration
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

# MySql Connection
def getDbConnection():
	ipAddress = '127.0.0.1'
	userId = 'root'
	userPassword = 'root'
	dataBase = 'stock'
	connection = pymysql.connect(host = ipAddress, user = userId, password = userPassword, db = dataBase, charset = 'utf8')
	return (connection)

# executeSqlCommand
def insertStockData(connection, sqlCommand, stockCode, instStockChart):	
	# Get Cursor
	cursor = connection.cursor()
	numberOfData = instStockChart.GetHeaderValue(3)
	for i in range(numberOfData) :
		dateOfData = instStockChart.GetDataValue(0, i);
		marketPrice = instStockChart.GetDataValue(1, i);
		highestPrice = instStockChart.GetDataValue(2, i);
		lowestPrice = instStockChart.GetDataValue(3, i);
		closingPrice = instStockChart.GetDataValue(4, i);
		netChange = instStockChart.GetDataValue(5, i);
		volume = instStockChart.GetDataValue(6, i); 
		cursor.execute(sqlCommand, (stockCode, dateOfData, marketPrice, highestPrice, lowestPrice,
			closingPrice, netChange, volume))
	# Commit
	print("insert stock data ", companyName)
	connection.commit()
	cursor.close()

# get stock header of stockCode
def getStockHeader(stockCode):
	instMarketEye = win32com.client.Dispatch("CpSysDib.MarketEye")
	instMarketEye.SetInputValue(1, stockCode)		# Set Stock Code
	instMarketEye.SetInputValue(0, 67)				# Set What to get (67 : PER)
	instMarketEye.BlockRequest()
	return instMarketEye

# update Stock Header
def updateStockHeader(connection, stockCode, companyName, instMarketEye):
	# Set Data
	per = instMarketEye.GetDataValue(0, 0)
	# Check Already Inserted Or Not
	sqlCommandCheckAlreadyInserted = "SELECT StockCode FROM StockHeader WHERE StockCode = %s"
	cursor = connection.cursor()
	cursor.execute(sqlCommandCheckAlreadyInserted, stockCode)
	connection.commit()
	row = cursor.fetchone()
	# Update Or Insert
	if row is not None :
		print("update stock header ", companyName)
		# update
		sqlCommandUpdateStockHeader = "UPDATE StockHeader SET CompanyName = %s, PER = %s WHERE StockCode = %s"
		cursor.execute(sqlCommandUpdateStockHeader, (companyName, per, stockCode))
	else :
		print("insert stock header ", companyName)
		# insert
		sqlCommandInsertStockHeader = """INSERT INTO StockHeader(StockCode, CompanyName, PER)
									VALUES (%s, %s, %s)"""
		cursor.execute(sqlCommandInsertStockHeader, (stockCode, companyName, per))
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



# 1. call connection check
connectionCheck()
 
# 2. Get All Comapany Info
stockCodeList = getAllStockCode()

# 3. getDbConnection
connection = getDbConnection()

for stockCode in stockCodeList.keys() :
	companyName = stockCodeList[stockCode]
	print("key : %s, value : %s" % (stockCode, companyName))
	# 4. update Stock Header
	# 4.1 get Stock Header
	instMarketEye = getStockHeader(stockCode)
	# 4.2 update Stock Header
	updateStockHeader(connection,  stockCode, companyName, instMarketEye)
	# 5. get Stock Data
	instStockChart = getStockData(stockCode)
	# 6. Make Sql Command
	sqlCommandInsertStockData = """INSERT INTO StockData(StockCode, DateOfData, MarketPrice, HighestPrice, LowestPrice, ClosingPrice, NetChange, Volume)
				VALUES (%s, %s, %s, %s, %s, %s, %s, %s)"""
	# 7. Execute Sql Command
	insertStockData(connection, sqlCommandInsertStockData, stockCode, instStockChart)
	break

# 8. close connection
connection.close()
