import cv2 
import numpy as np
import win32com.client

def onChange(x):
	pass
	

def mouse_click(event, x, y, flags, param):
	global start_ppt

	if event == cv2.EVENT_LBUTTONDOWN:
		start_ppt = not start_ppt
		print("Start prezentacja"+str(start_ppt))
	elif event == cv2.EVENT_LBUTTONUP:
		pass

cap = cv2.VideoCapture(0)

cv2.namedWindow("Tracking")
cv2.createTrackbar("LH", "Tracking", 0, 255, onChange)
cv2.createTrackbar("LS", "Tracking", 0, 255, onChange)
cv2.createTrackbar("LV", "Tracking", 0, 255, onChange)
cv2.createTrackbar("UH", "Tracking", 0, 255, onChange)
cv2.createTrackbar("US", "Tracking", 0, 255, onChange)
cv2.createTrackbar("UV", "Tracking", 0, 255, onChange)



start_ppt = False
windowName = 'res'
cv2.namedWindow(windowName)
cv2.setMouseCallback(windowName, mouse_click)
#powerpoint
app = win32com.client.Dispatch("PowerPoint.Application")
app.Visible = True
presentation = app.Presentations.Open('D:\Python\gestures\prezentacja2.pptx', WithWindow =1 ) #path to the presentation

##############
gestures = ['start','stop','next','prev']

start = cv2.imread('D:\Python\gestures\start.jpg') #path to image with gesture's mask
start = cv2.resize(start,(600,450))

stop = cv2.imread('D:\Python\gestures\stop.jpg')
stop = cv2.resize(stop,(600,450))

prev = cv2.imread('D:\Python\gestures\prev.jpg')
prev = cv2.resize(prev,(600,450))

next = cv2.imread('D:\Python\gestures\gonext.jpg')
next = cv2.resize(next,(600,450))




cnt = 0
test = 0
previous = " "
while True:

	
	_, frame = cap.read()


	fSize = cv2.resize(frame, (600,450))

	hsv = cv2.cvtColor(frame, cv2.COLOR_BGR2HSV)

	l_h = cv2.getTrackbarPos("LH","Tracking")
	l_s = cv2.getTrackbarPos("LS","Tracking")
	l_v = cv2.getTrackbarPos("LV","Tracking")

	u_h = cv2.getTrackbarPos("UH","Tracking")
	u_s = cv2.getTrackbarPos("US","Tracking")
	u_v = cv2.getTrackbarPos("UV","Tracking")


	l_b = np.array([l_h, l_s, l_v])
	u_b = np.array([u_h, u_s, u_v])

	mask = cv2.inRange(hsv,l_b, u_b)
	mSize = cv2.resize(mask,(600,450))

	res = cv2.bitwise_and(frame, frame, mask=mask)
	rSize= cv2.resize(res,(600,450))
	
	hand = res
	hSize = cv2.resize(hand,(600,450))	
	

	masks = []
	results = []
	

	
	masks.append(cv2.cvtColor(start, cv2.COLOR_BGR2GRAY))
	masks.append(cv2.cvtColor(stop, cv2.COLOR_BGR2GRAY))
	masks.append(cv2.cvtColor(next, cv2.COLOR_BGR2GRAY))
	masks.append(cv2.cvtColor(prev, cv2.COLOR_BGR2GRAY))
	
	results.append(cv2.bitwise_and(start, start, mask=masks[gestures.index('start')]))
	results.append(cv2.bitwise_and(stop, stop, mask=masks[gestures.index('stop')]))
	results.append(cv2.bitwise_and(next, next, mask=masks[gestures.index('next')]))
	results.append(cv2.bitwise_and(prev, prev, mask=masks[gestures.index('prev')]))
	

	
	try:


		
		contours, hierarchy = cv2.findContours(mSize, cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_NONE)
		biggest_input = max(contours, key= cv2.contourArea)
		x,y,w,h = cv2.boundingRect(biggest_input)
		cv2.rectangle(hand,(x,y),(x+w,x+h),(0,255,0),2)
		hSize = cv2.resize(hand,(600,450))
		
	
		matching = []
		
		for m, r in zip (masks, results):
			contours, hierarchy = cv2.findContours(m, cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_NONE)
			biggest_sample = max(contours, key= cv2.contourArea)
			x,y,w,h = cv2.boundingRect(biggest_sample)
			cv2.rectangle(r,(x,y),(x+w,x+h),(0,255,0),2)
	
			matching.append(cv2.matchShapes(biggest_sample,biggest_input,1,0.0))
		
		
		detected_gesture = gestures[matching.index(min(matching))]
		value = min(matching)
		
		#threshold to avoid detecting noice
		if detected_gesture == 'start' or detected_gesture == 'stop':
			if value > 0.2:
				detected_gesture = 'none'
		elif detected_gesture == 'next':
			if value > 0.3:
				detected_gesture = 'none'
		elif detected_gesture == 'prev':
			if value > 1:
				detected_gesture = 'none'
		
		print(matching)
		print(detected_gesture)
		
	
		
		m = cv2.moments(mSize)
		cX = int(m["m10"]/m["m00"])
		cY = int(m["m01"]/m["m00"])
		cv2.circle(fSize, (cX, cY), 5,(255,255,255), -1)

	except Exception as e: 
		test += 1
		
	
	if 	start_ppt == True:
		if previous == " ":
			previous = detected_gesture
		elif previous == detected_gesture:
			previous = detected_gesture
			cnt += 1
		else:
			cnt = 0
			previous = " "
		
		if cnt == 15:
			try:
				if detected_gesture == 'start':
					print("START")
					presentation.SlideShowSettings.Run()
				elif detected_gesture == 'stop':		
					print("STOP")			
					presentation.SlideShowWindow.View.Exit()
				elif detected_gesture == 'next':
					print("NEXT")
					presentation.SlideShowWindow.View.Next()
				elif detected_gesture == 'prev':
					print("PREV")
					presentation.SlideShowWindow.View.Previous()
				cnt = 0
			except Exception as e: 
				print(e)
	


	cv2.imshow("frame", fSize)
	#cv2.imshow("mask", mSize)
	cv2.imshow("res", rSize)
	#cv2.imshow("hand",hSize)
	
	

    
	key = cv2.waitKey(1)


cap.release()
cv2.destroyAllWindows()
