'''
Created on November 20, 2010
@author: Dr. Rainer Hessmer

Modified by Ivan Perez Olivera (ivan.perezolivera@hotmail.com) to find the COM port automatically given a dictionary of devices
'''

import threading
import serial
from serial.tools import list_ports
from io import StringIO
import time
import re

def LookForDevices(devices):
	scaleDevice = None
	scannerDevice = None 
	ports = list_ports.comports()
	
	# look for scale
	for k in devices["Scale"].keys():
		for c in ports:
			if re.findall(k,c.hwid):
				scaleDevice = ("Scale,{},{}".format(c.device,k))
				break

	#look for scanner
	for k in devices["Scanner"].keys():
		for c in ports:
			if re.findall(k,c.hwid):
				scannerDevice = ("Scanner,{},{}".format(c.device,k))
				break
	return( (scaleDevice,scannerDevice) )

def _OnLineReceived(line):
	print(line)

class SerialDataGateway(object):
	'''
	Helper class for receiving lines from a serial port
	'''

	def __init__(self, port="COM8", baudrate=9600,bytesSize = 8,stopBits = 1, lineHandler = _OnLineReceived):
		'''
		Initializes the receiver class. 
		port: The serial port to listen to.
		receivedLineHandler: The function to call when a line was received.
		'''
		self._Port = port
		self._Baudrate = baudrate
		self._BytesSize = bytesSize
		self._StopBits = stopBits
		self._ReceivedLineHandler = lineHandler
		self._KeepRunning = False
		self._bytesReceived = 0

	def Start(self):
		try:
			self._Serial = serial.Serial(port = self._Port, baudrate = self._Baudrate, bytesize = self._BytesSize,stopbits = self._StopBits, timeout = 1)
			self._KeepRunning = True
			self._ReceiverThread = threading.Thread(target=self._Listen)
			self._ReceiverThread.setDaemon(True)
			self._ReceiverThread.start()
		except:
			print("Port not found")

	def Stop(self):
		print("Stopping serial gateway")
		self._KeepRunning = False
		time.sleep(.1)
		self._Serial.close()

	def _Listen(self):
		stringIO = StringIO()
		while self._KeepRunning:
			data = self._Serial.read()
			data = data.decode("ASCII")
			if data == '\r' or data == '':
				continue
			if data == '\n' and self._bytesReceived > 0:
				self._ReceivedLineHandler(stringIO.getvalue())
				stringIO.close()
				stringIO = StringIO()
				self._bytesReceived = 0
			if data == '\n' and self._bytesReceived == 0:
				stringIO.close()
				stringIO = StringIO()
			else:
				stringIO.write(data)
				self._bytesReceived += 1

	def Write(self, data):
		info = "Writing to serial port: %s" %data
		print(info)
		self._Serial.write(data)

	if __name__ == '__main__':
		dataReceiver = SerialDataGateway("COM21",  9600)
		dataReceiver.Start()
		input("Hit <Enter> to end.")
		dataReceiver.Stop()
		print("Execution stopped")