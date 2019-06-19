import pymysql
import mysql.connector
import json
import sys
from mysql.connector import errorcode, locales
import xlrd
from xlutils.copy import copy as xl_copy
import re
import pythoncom
from win32com.client.gencache import EnsureDispatch
 
try:
	commandKey = sys.argv[1]
	
	if commandKey == "ConnectDB":
		try:
			if len(sys.argv) == 6:
				# opening a database connection
				host = sys.argv[2]
				user = sys.argv[3]
				password = sys.argv[4]
				DBName = sys.argv[5]
				db = mysql.connector.connect(user=user, password=password,host=host,database=DBName)
				data = {}
				data['status_code'] = "200 OK"
				data['status'] = "Connected successful"
				print(json.dumps(data))
				
			else:
				self.fail("locales.eng.client_error could not be imported")
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide proper parameters"
				print(json.dumps(data))
				sys.exit()		
				
			
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	
			
			
	elif commandKey == "DisconnectDB":
		try:
			if len(sys.argv) == 6:
				# opening a database connection
				host = sys.argv[2]
				user = sys.argv[3]
				password = sys.argv[4]
				DBName = sys.argv[5]
				db = mysql.connector.connect(user=user, password=password,host=host,database=DBName)
				db.close()
				data = {}
				data['status_code'] = "200 OK"
				data['status'] = "Disconnected successful"
				print(json.dumps(data))
				
			else:
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide proper parameters"
				print(json.dumps(data))
				sys.exit()
				
			
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	
			
					
	elif commandKey == "ExecuteQuery":
		try:
			if len(sys.argv) == 9:
				# opening a database connection
				host = sys.argv[2]
				user = sys.argv[3]
				password = sys.argv[4]
				DBName = sys.argv[5]
				sql = sys.argv[6]
				wb_name = sys.argv[7]
				wk_sheet = sys.argv[8]
				db = mysql.connector.connect(user=user, password=password,host=host,database=DBName)

				sql1 = sql.lower()
				words = sql1.split()
				
				if "select" in words:
					#cleaning the existing data in excel sheet
					rb = xlrd.open_workbook(wb_name)
					wb = xl_copy(rb)
					Sheet1 = wb.add_sheet(wk_sheet)
					#wb.save(wb_name)
					sheet = wb.get_sheet(wk_sheet)	 
						 
					text_query = sql.lower()
					text_to_edit = re.search('select (.+?) from', str(text_query))
					if text_to_edit:
						found_to_edit = text_to_edit.group(1)
					
					if found_to_edit == "*":
						cur1 = db.cursor()
						# Use all the SQL you like
						#text_query1 = sql.lower()
						#text_to_edit1 = text_query1.split("from",1)[1] 
						text_query1 = sql.lower()
						list_of_words = text_query1.split()
						text_to_edit1 = list_of_words[list_of_words.index("from") + 1]
							
						cur1.execute("SHOW COLUMNS FROM " + text_to_edit1)

						#print(cur1.fetchall())
						db_data1 = cur1.fetchall()
						db_data2 = db_data1
						db_data3 = db_data1
						k = len(db_data1)

						# print all the cells of the row to excel sheet

						rowNum = 0
						colNum = 0 #keep track of columns
						for i in range(0, k):
							#value = sh.cell_value(k,i)
							sheet.write(rowNum, colNum, db_data3[i][0])
							colNum = colNum + 1
							#k = k+1
						#wb.save(wb_name)

						# you must create a Cursor object. It will let
						# you execute all the queries you need
						cur1.close()
								
					else:
						found_to_edit_arr = found_to_edit.split(",")
						found_to_edit_arr1 = list(found_to_edit_arr)
						j = len(found_to_edit_arr1)
						
						rowNum = 0
						colNum = 0 #keep track of columns

						for i in range(0, j):
							sheet.write(rowNum, colNum, found_to_edit_arr1[i])
							colNum = colNum + 1
						j = j + 1
						#wb.save(wb_name)

					# prepare a cursor object using cursor() method
					cursor = db.cursor()
					
					
					
					cursor.execute(sql)
					resul1 = cursor.fetchall()
					resul2 = resul1
					
					#cursor.execute("SELECT * FROM country")
					#print(cursor.description)
					#print(" Test")
						
					#print(cursor.fetchall())

					
					db_data1 = resul2
					
					db_data2 = db_data1
					
					db_data3 = db_data1
					k = len(db_data1[0])
					
					'''
					for row in db_data3:
						print(str(row))
					'''
					#db.commit()
					
					rowNum = 1
					for row in db_data2:
						#colNum = 0 #keep track of columns
						for i in range(0, k):
							colNum = i
							#excel_list21= sheet.cell(r,excel_lists2.index(excel_list2)).value
							sheet.write(rowNum, colNum, row[i])
							#colNum = colNum + 1
						rowNum = rowNum + 1	
						#k = k+1
					wb.save(wb_name)	
					cursor.close()
					data = {}
					data['status_code'] = "200 OK"
					data['data'] = str(resul1)
					if len(resul1) > 0:
						print(json.dumps(data))	
					#conn.close()

					# disconnect from server
					db.close()
					
				else:			
					data = {}
					data['status_code'] = "401"
					data['status'] = "Please provide only select command"
					print(json.dumps(data))
					sys.exit()
		
			else:
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide proper parameters"
				print(json.dumps(data))
				sys.exit()
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass		
	
	
	elif commandKey == "ExecuteNonQuery":
		try:
			# opening a database connection
			if len(sys.argv) > 6:
				host = sys.argv[2]
				user = sys.argv[3]
				password = sys.argv[4]
				DBName = sys.argv[5]
				db = mysql.connector.connect(user=user, password=password,host=host,database=DBName)
				sql = sys.argv[6]
				sql1 = sql.lower()
				words = sql1.split()
				
			elif len(sys.argv) == 6:
				host = sys.argv[2]
				user = sys.argv[3]
				password = sys.argv[4]
				sql = sys.argv[5]
				sql1 = sql.lower()
				words = sql1.split()
				db = mysql.connector.connect(user=user, password=password,host=host)
				
			else:
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide proper parameters"
				print(json.dumps(data))
				sys.exit()
				
			# prepare a cursor object using cursor() method
			cursor = db.cursor()

			# Drop table if it already exist using execute() method.
			#cursor.execute("SELECT * from )

			# Create table as per requirement
			#sql = sys.argv[6]
			
			#print(sql)
			if "select" in words:
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide proper query, other than select"
				print(json.dumps(data))
				pass
			else:
				result = cursor.execute(sql)
				#print("Hi")
				#print(cursor.rowcount)
				db.commit()
				data = {}
				data['status_code'] = "200 OK"
				if cursor.rowcount == 1:
					data['status'] = str(cursor.rowcount) +" row affected"
				else:	
					data['status'] = str(cursor.rowcount) +" rows affected"
				print(json.dumps(data))
				cursor.close()
				#conn.close()

				# disconnect from server
				db.close()
			
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
			
	elif commandKey == "InsertDB":
		try:
			if len(sys.argv) == 7:
				# opening a database connection
				host = sys.argv[2]
				user = sys.argv[3]
				password = sys.argv[4]
				DBName = sys.argv[5]
				db = mysql.connector.connect(user=user, password=password,host=host,database=DBName)

				# prepare a cursor object using cursor() method
				cursor = db.cursor()

				# Drop table if it already exist using execute() method.
				#cursor.execute("SELECT * from )

				# Create table as per requirement
				sql = sys.argv[6]
				sql1 = sql.lower()
				words = sql1.split()
				if "insert" in words:
				
					#print(sql)

					cursor.execute(sql)
					db.commit()
					data = {}
					data['status_code'] = "200 OK"
					data['status'] = "Task completed successfully"
					print(json.dumps(data))
					cursor.close()
					#conn.close()

					# disconnect from server
					db.close()
					
				else:			
					data = {}
					data['status_code'] = "401"
					data['status'] = "Please provide only insert command"
					print(json.dumps(data))
					sys.exit()
			
			else:
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide proper parameters"
				print(json.dumps(data))
				sys.exit()
			
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass		
			
			
	elif commandKey == "ImportExceltoDB":
		try:		
			# Establish a MySQL connection
			host = sys.argv[2]
			user = sys.argv[3]
			password = sys.argv[4]
			wb_name = sys.argv[5]
			wk_sheet = sys.argv[6]
			DBName = sys.argv[7]
			table_name = sys.argv[8]
			excel_list = sys.argv[9]
			db_list = sys.argv[10]


			# Open the workbook and define the worksheet
			book = xlrd.open_workbook(wb_name)
			sheet = book.sheet_by_name(wk_sheet)

			db = mysql.connector.connect(user=user, password=password,host=host,database=DBName)

			# Get the cursor, which is used to traverse the database, line by line
			cursor = db.cursor()

			#splitting excel inputs
			excel_list1 = excel_list.split(',')
			excel_lists2 = list(excel_list1)

			#splitting database names
			db_list1 = db_list.split(',')
			db_lists2 = list(db_list1)

			def findexcel_column(excel_list2):
				book = xlrd.open_workbook(wb_name)
				sh = book.sheet_by_name(wk_sheet)
				for row in range(sh.nrows):
					for column in range(sh.ncols):
						if excel_list2 == sh.cell(row, column).value:  
							return column


			table_names_list = []

			for db_list in db_lists2:
				table_names_list.append("%s")
			# Create the INSERT INTO sql query
			db_str1 = table_name
			db_str_out2 = ','.join(str(e) for e in db_lists2)
			db_str_val2 = ','.join(str(e) for e in table_names_list)
			excel_str_val2 = ','.join(str(e) for e in excel_lists2)


			query = "INSERT INTO " + db_str1+" (" + db_str_out2 + ") VALUES ("+db_str_val2+")"

			# Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers

			for r in range(1, sheet.nrows):
				excel_lists2_arr = []
				for excel_list2 in excel_lists2:
					findexcel_column_num = findexcel_column(excel_list2)
					#excel_list21= sheet.cell(r,excel_lists2.index(excel_list2)).value
					excel_list21= sheet.cell(r,int(findexcel_column_num)).value
					excel_list21 =  excel_list21
					excel_lists2_arr.append(excel_list21)
				
				# Execute sql Query
				cursor.execute(query, excel_lists2_arr)

			# Close the cursor
			cursor.close()

			# Commit the transaction
			db.commit()
			data = {}
			data['status_code'] = "200 OK"
			data['status'] = "Data imported successfully"
			print(json.dumps(data))
			# Close the database connection
			db.close()

			
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	
			
	
	elif commandKey == "ExportDBTabletoExcel":
		try:		
			host = sys.argv[2]
			user = sys.argv[3]
			password = sys.argv[4]
			wb_name = sys.argv[5]
			wk_sheet = sys.argv[6]
			DBName = sys.argv[7]
			table_name = sys.argv[8]
			#excel_list = sys.argv[8]
			#db_list = sys.argv[9]
			# Open the workbook and define the worksheet
			rb = xlrd.open_workbook(wb_name)
			wb = xl_copy(rb)
			sheet = wb.get_sheet(wk_sheet)

			#Database connection
			db = mysql.connector.connect(user=user, password=password,host=host,database=DBName)

			cur1 = db.cursor()
			# Use all the SQL you like
			cur1.execute("SHOW COLUMNS FROM " + table_name)

			#print(cur1.fetchall())
			db_data1 = cur1.fetchall()
			db_data2 = db_data1
			db_data3 = db_data1
			k = len(db_data1)

			# print all the cells of the row to excel sheet

			rowNum = 0
			colNum = 0 #keep track of columns
			for i in range(0, k):
				#value = sh.cell_value(k,i)
				sheet.write(rowNum, colNum, db_data3[i][0])
				colNum = colNum + 1
				#k = k+1
			wb.save(wb_name)

			# you must create a Cursor object. It will let
			# you execute all the queries you need
			cur1.close()

			cur = db.cursor()
			# Use all the SQL you like
			cur.execute("SELECT * FROM " + table_name)
			#k = 0
			#sh = rb.sheet_by_name(wk_sheet)
			#total_cols = sh.ncols
			# print all the cells of the row to excel sheet
			rowNum = 1
			for row in cur.fetchall():
				colNum = 0 #keep track of columns
				for i in range(0, k):
					sheet.write(rowNum, colNum, row[i])
					colNum = colNum + 1
				rowNum = rowNum + 1	
				#k = k+1
			wb.save(wb_name)	
			data = {}
			data['status_code'] = "200 OK"
			data['status'] = "Data exported successfully"
			print(json.dumps(data))
				#sheet.write(rowNum, colNum, row) # row, column, value
			cur.close()
			# Close the database connection
			db.close()
						
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	
			
	
	elif commandKey == "ExportDBTableColumnstoExcel":
		try:		
			host = sys.argv[2]
			user = sys.argv[3]
			password = sys.argv[4]
			wb_name = sys.argv[5]
			wk_sheet = sys.argv[6]
			DBName = sys.argv[7]
			table_name = sys.argv[8]
			db_list = sys.argv[9]
			excel_list = sys.argv[10]
			# Open the workbook and define the worksheet
			rb = xlrd.open_workbook(wb_name)
			wb = xl_copy(rb)
			sheet = wb.get_sheet(wk_sheet)
			db = mysql.connector.connect(user=user, password=password,host=host,database=DBName)
			#splitting excel inputs
			
			def findexcel_column(excel_list2):
				book = xlrd.open_workbook(wb_name)
				sh = book.sheet_by_name(wk_sheet)
				for row in range(sh.nrows):
					for column in range(sh.ncols):
						if excel_list2 == sh.cell(row, column).value:  
							return column
								
								
			excel_list1 = excel_list.split(',')
			excel_lists2 = list(excel_list1)
			
			#splitting database names
			db_list1 = db_list.split(',')
			db_lists2 = list(db_list1)
			
			k = len(excel_lists2)
			if len(excel_lists2) == len(db_lists2):
				# print all the cells of the row to excel sheet
								
				x = 0				
				rowNum = 0
				#colNum = 0 #keep track of columns
				for i in excel_lists2:
					#value = sh.cell_value(k,i)
					colNum = findexcel_column(i)
					sheet.write(rowNum, int(colNum), i)
					rowNum = rowNum + 1
					#k = k+1
				wb.save(wb_name)

				cur = db.cursor()
				# Use all the SQL you like
				cur.execute("SELECT " + db_list + " FROM " + table_name)
				#k = 0
				#sh = rb.sheet_by_name(wk_sheet)
				#total_cols = sh.ncols
				# print all the cells of the row to excel sheet
				
				rowNum = 1
				for row in cur.fetchall():
					#colNum = 0 #keep track of columns
					j = 0
					for i in range(0, k):
						colNum = findexcel_column(excel_lists2[j])
						#excel_list21= sheet.cell(r,excel_lists2.index(excel_list2)).value
						sheet.write(rowNum, colNum, row[i])
						#colNum = colNum + 1
						j = j + 1
					rowNum = rowNum + 1	
					#k = k+1
				wb.save(wb_name)	
				data = {}
				data['status_code'] = "200 OK"
				data['status'] = "Data exported successfully"
				print(json.dumps(data))
					#sheet.write(rowNum, colNum, row) # row, column, value
				cur.close()
				# Close the database connection
				db.close()
				
			else:	
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide proper field names"
				print(json.dumps(data))
				pass
			
						
		except OSError as e:
				data = {}
				data['status_code'] = "401"
				data['status'] = str(e)
				print(json.dumps(data))
				pass	
	
	
	else:
		print("Please enter proper Command Key")
	
	
except Exception as e:
	data = {}
	data['status_code'] = "401"
	data['status'] = str(e) 
	#data = json.dumps(data)
	print(json.dumps(data))
	pass	