import mysql.connector
from mysql.connector import errorcode
import pandas as pd



def populate_data(array):
	print array[0][0]
	print array[1][0]


def iter_row(cursor,size=2):
	while True:
		rows = cursor.fetchmany(size)
		if not rows:
			break
		for row in rows:
			yield row
	pass

def Query_data(cnx,array):

	try:
		cursor = cnx.cursor()
		cursor.execute("SELECT * FROM entries")
		# row = cursor.fetchone()

		for row in iter_row(cursor,2):
			array.append(row)
			print(row)

			
	except Error as e:
		print(e)
	finally:
		cursor.close()


def connect(array):
	
	try:
	  cnx = mysql.connector.connect(user='root',password='password@123',host='127.0.0.1',
	                                database='invescoapp_development')
	except mysql.connector.Error as err:
	  if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
	    print("Something is wrong with your user name or password")
	  elif err.errno == errorcode.ER_BAD_DB_ERROR:
	    print("Database does not exist")
	  else:
	    print(err)
	else:
		print "Success!"
		if cnx.is_connected():
			print('Connected to  MYSQL database')
			Query_data(cnx,array)

	finally : 
		print "Closing connection."
		cnx.close()




if __name__ == '__main__':
    array = []
    connect(array)
    populate_data(array)