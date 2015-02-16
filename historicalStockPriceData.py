from yahoo_finance import Share
import xlsxwriter
import webbrowser

class historicalData():

	def __init__(self):
		stockSymbol = raw_input("Enter Stock Symbol: ")
		startDate = raw_input("Enter Start Date (year-mm-dd): ")
		endDate = raw_input("Enter End Date (year-mm-dd): ")
		stockSymbol = stockSymbol.upper()
		self.stockSymbol = stockSymbol
		self.startDate = startDate
		self.endDate = endDate
		stock = Share(stockSymbol)
		global companyStockSymbol
		companyStockSymbol = stockSymbol
		global historical 
		historical = stock.get_historical(startDate, endDate)

	def getHistoricalData(self):
		#filename 
		filename = companyStockSymbol + '_historicalPriceData.xlsx'
		# Create a workbook and add a worksheet.
 		workbook = xlsxwriter.Workbook(filename)
 		worksheet = workbook.add_worksheet()
 		# Add a bold format to use to highlight cells.
 		bold = workbook.add_format({'bold': True})
 		# Add a number format for cells with money.
		money = workbook.add_format({'num_format': '$#,##0.00'})
		#Add a number format for cells with vol
		vol = workbook.add_format({'num_format': '#,##0'})
		#Format cell width 
		worksheet.set_column('A:G', 15)
		# Write some data headers.
 		worksheet.write('A1', 'Volume', bold)
 		worksheet.write('B1', 'Adj_close', bold)
 		worksheet.write('C1', 'High', bold)
 		worksheet.write('D1', 'Low', bold)
 		worksheet.write('E1', 'Date', bold)
 		worksheet.write('F1', 'Close', bold)
 		worksheet.write('G1', 'Open', bold)
 		# Start from the first cell below the headers.
 		row = 1
 		col = 0


 		# Iterate over the data and write it out row by row.
 		for item in historical:
 			count = 0
			for v in item.itervalues():
				count += 1
				if count == 1:
					v = float(v)
					worksheet.write_number(row, col, v, vol)
				elif count == 2:
					continue
				elif count > 2 and count < 6:
					v = float(v)
					worksheet.write_number(row, col, v, money)
				elif count == 6:
					worksheet.write(row, col, v)
				elif count == 7 or count == 8:
					v = float(v)
					worksheet.write_number(row, col, v, money)
				col += 1
			row += 1
			col = 0
     	#Close workbook
		workbook.close()
		webbrowser.open(filename)

stockY = historicalData()
stockY.getHistoricalData()