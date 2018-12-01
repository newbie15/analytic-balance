import serial
import time
import win32api
import win32com.client

timbangan = serial.Serial (port = 'COM4',baudrate = 1200,parity = serial.PARITY_ODD,stopbits = serial.STOPBITS_ONE,bytesize = serial.SEVENBITS)

timbangan.xonoff = True
timbangan.dsrdtr = True
timbangan.rtscts = True

def main ():
	while True:
		data = timbangan.readline ()
		if data != "":
			print data
			databaru = data.replace ('N','') .replace (' ','') .replace ('+','') .replace ('g','')
			databaru = databaru.replace ('\n','') .replace ('\r','')
			# print "data yang terbaca", databaru

			shell = win32com.client.Dispatch ("WScript.Shell")

			shell.SendKeys (databaru)
			win32api.Sleep (500)

			# shell.SendKeys ("{ENTER}")
			# win32api.Sleep (2500)

		pass

	pass

if __name__ == '__main__':
	main()