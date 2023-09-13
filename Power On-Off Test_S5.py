#==============================================================================
#     日期           人員           版本      項目
#     ----------     ----------     -----     ---------------------------------
#     2023/08/31     TH Chen        1.0       初版本
#
#==============================================================================
import traceback
import sys
import os
import win32api
import win32gui
import subprocess
import pygetwindow as getwindow
import pyautogui
import datetime
import time
import requests
import urllib3
import uuid
import pywifi
from pywifi import const
import bluetooth


Config_File="Config.txt"
Log_File="HP_Power-ON-OFF_Log.txt"


try:
	#Config.txt讀檔
	start_time=datetime.datetime.now()
	Config_Setting=open(Config_File, mode='r', encoding='utf-8')

	#Config.txt
	#第1行: AP名稱
	#第2行: BT MAC Address
	#第3行: 執行的開始值
	#第4行: 執行的結束值
	#第5行: 連接AP時, 等待秒數
	#第6行: 連接BT時, 等待秒數
	#第7行: WIFI與BT執行秒數
	#第8行: 關機等待時間
	issid=Config_Setting.readline()
	issid=issid.replace('\n','')
	ibt=Config_Setting.readline()
	ibt=ibt.replace('\n','')
	istart=Config_Setting.readline()
	istart=int(istart.replace('\n',''))
	iend=Config_Setting.readline()
	iend=int(iend.replace('\n',''))
	iloop1=Config_Setting.readline()
	iloop1=int(iloop1.replace('\n',''))
	iloop2=Config_Setting.readline()
	iloop2=int(iloop2.replace('\n',''))
	irun=Config_Setting.readline()
	irun=int(irun.replace('\n',''))
	ishutdown=Config_Setting.readline()
	ishutdown=int(ishutdown.replace('\n',''))
	Config_Setting.close()


	# 如果開始次數小於等於結束次數
	if (istart <= iend):
		#Log開檔 (如果存在就增加至檔案最後, 如果不存在就新增)
		Log_File=open(Log_File, mode='a+', encoding='utf-8')
		iap_result=False

		os.system("cls")
		idatetime=datetime.datetime.now()
		print("\033[32m[",idatetime,"] \033[37m","Loop: ",istart,sep='',end='\n')
		print("[",idatetime,"] Loop: ",istart,sep='',end='\n',file=Log_File)

		# 查詢是否有 Wi-Fi interface
		# cmd指令: netsh interface show interface
		iwifi_interface=0
		cmd_all_output=subprocess.check_output(["netsh","interface","show","interface"],stderr=subprocess.STDOUT,universal_newlines=True)
		cmd_lines=cmd_all_output.split("\n")
		for cmd_line in cmd_lines:
			if "Wi-Fi" in cmd_line:
				iwifi_interface=1
				break

		# 如果有 Wi-Fi interface
		if(iwifi_interface > 0):
			iwifi = pywifi.PyWiFi()
			iinterface = iwifi.interfaces()[0]

			# 斷線
			iinterface.disconnect()
			time.sleep(1)

			# 取得Wi-Fi MAC address
			wifi_mac_address=uuid.UUID(int=uuid.getnode()).hex[-12:]
			wifi_mac_address=wifi_mac_address[:-1]+format(int(wifi_mac_address[-1],16)-1,'x')
			wifi_mac_address=wifi_mac_address.upper()
			wifi_mac_address=':'.join(wifi_mac_address[i:i+2] for i in range(0,len(wifi_mac_address),2))
			idatetime=datetime.datetime.now()
			print("\033[32m[",idatetime,"] \033[37m","Wi-Fi interface was found (",wifi_mac_address,")                    ",sep='',end='\n')
			print("[",idatetime,"]"," Wi-Fi interface was found (",wifi_mac_address,")                    ",sep='',end='\n',file=Log_File)

			# 搜尋連線AP是否存在
			idatetime=datetime.datetime.now()
			print("\033[32m[",idatetime,"] \033[37m","AP(",issid,") is searching...waiting ",sep='',end='\r')
			iinterface.scan()

			# 搜尋AP等待時間
			itemp1=0
			itemp2=round(iloop1/2)
			while (itemp1 < itemp2):
				idatetime=datetime.datetime.now()
				print("\033[32m[",idatetime,"] \033[37m","AP(",issid,") is searching...waiting ",itemp2,"          ",sep='',end='\r')
				time.sleep(1)
				itemp2=itemp2-1
			iscan_ap = iinterface.scan_results()
			for ap_name in iscan_ap:
				if (ap_name.ssid.upper() == issid.upper()):
					iap_result=True
					break

			# 如果AP存在, 則連線
			if (iap_result == True):

				idatetime=datetime.datetime.now()
				print("\033[32m[",idatetime,"] \033[37m","AP(",issid,") was found                                        ",sep='',end='\n')
				print("[",idatetime,"] ","AP(",issid,") was found",sep='',end='\n',file=Log_File)

    			# 連線指定AP (如:DQA_2.4G)
				connect_ap = pywifi.Profile()
				connect_ap.ssid = issid
				ap = iinterface.add_network_profile(connect_ap)
				iinterface.connect(ap)

				# 連線AP等待時間
				itemp1=0
				itemp2=round(iloop1/2)
				while (itemp1 < itemp2):
					idatetime=datetime.datetime.now()
					print("\033[32m[",idatetime,"] \033[37m","AP(",issid,") is connecting...waiting ",itemp2,"          ",sep='',end='\r')
					time.sleep(1)
					itemp2=itemp2-1

				# 判斷AP連線是否成功
				idatetime=datetime.datetime.now()
				if (iinterface.status() == pywifi.const.IFACE_CONNECTED):
					print("\033[32m[",idatetime,"] \033[37m","AP(",issid,") connection successful","          ",sep='',end='\n')
					print("[",idatetime,"] AP(",issid,") connection successful",sep='',end='\n',file=Log_File)
				else:
					print("\033[32m[",idatetime,"] \033[91m","AP(",issid,") connection failed","          ",sep='',end='\n')
					print("[",idatetime,"] AP(",issid,") connection failed",sep='',end='\n',file=Log_File)

			else:
				idatetime=datetime.datetime.now()
				print("\033[32m[",idatetime,"] \033[91m","AP(",issid,") does not exist","                    ",sep='',end='\n')
				print("[",idatetime,"] ","AP(",issid,") does not exist",sep='',end='\n',file=Log_File)


			# 取得BT MAC address
			bt_mac_address = bluetooth.read_local_bdaddr()
			time.sleep(1)
			idatetime=datetime.datetime.now()
			print("\033[32m[",idatetime,"] \033[37m","BT interface was found (",bt_mac_address[0],")                    ",sep='',end='\n')
			print("[",idatetime,"]"," BT interface was found (",bt_mac_address[0],")                    ",sep='',end='\n',file=Log_File)

			# 搜尋BT藍牙喇叭是否與電腦配對
			idatetime=datetime.datetime.now()
			print("\033[32m[",idatetime,"] \033[37m","BT(",ibt,") is checking pair information...",sep='',end='\r')
			ibtspeaker=bluetooth.discover_devices(duration=8, lookup_names=True,device_id=-1)
			time.sleep(1)
			ibt_result=False
			for bt_mac,bt_name in ibtspeaker:
				if (bt_mac.upper() == ibt.upper()):
					ibt_result=True
					break

			# 如果BT藍牙喇叭與電腦配對
			if (ibt_result == True):

				idatetime=datetime.datetime.now()
				print("\033[32m[",idatetime,"] \033[37m","BT(",ibt,") has paired with computer","                    ",sep='',end='\n')
				print("[",idatetime,"] ","BT(",ibt,") has paired with computer",sep='',end='\n',file=Log_File)

				# 搜尋BT藍牙喇叭是否開機可連線
				# cmd指令: btdiscovery -b 20:64:DE:85:8B:FE -s"%sc%"
				idatetime=datetime.datetime.now()
				print("\033[32m[",idatetime,"] \033[37m","BT(",ibt,") is detecting...","                    ",sep='',end='\r')
				ibt_result=False
				bt_check_01="btdiscovery"
				bt_check_01_1=["-b",ibt,"-s\"%sc%\""]
				cmd_all_output=subprocess.run([bt_check_01]+bt_check_01_1, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True, text=True)
				time.sleep(1)
				cmd_output=cmd_all_output.stdout.splitlines()
				if (len(cmd_output) > 1):
					idatetime=datetime.datetime.now()
					print("\033[32m[",idatetime,"] \033[37m","BT(",ibt,") detection successful","                    ",sep='',end='\n')
					print("[",idatetime,"] BT (",ibt,") detection successful",sep='',end='\n',file=Log_File)

					# 連線BT藍牙喇叭 (如:20:64:DE:85:8B:FE)
					# cmd指令: btcom -b 20:64:DE:85:8B:FE -r -s110b
					ibt_result=False
					bt_connect_01="btcom"
					bt_connect_01_1=["-b",ibt,"-r","-s110b"]
					cmd_all_output=subprocess.run([bt_connect_01]+bt_connect_01_1, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True, text=True)
					time.sleep(1)
					if (cmd_all_output.returncode == 0):

						# 連線BT藍牙喇叭 (如:20:64:DE:85:8B:FE)
						# cmd指令: btcom -b 20:64:DE:85:8B:FE -c -s110b
						bt_connect_02="btcom"
						bt_connect_02_1=["-b",ibt,"-c","-s110b"]
						cmd_all_output=subprocess.run([bt_connect_02]+bt_connect_02_1, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True, text=True)

		   				# 等待BT藍牙喇叭連線時間
						itemp1=0
						itemp2=iloop2
						while itemp1<itemp2:
							idatetime=datetime.datetime.now()
							print("\033[32m[",idatetime,"] \033[37mBT(",ibt,") is connecting...waiting ",itemp2,"          ",sep='',end='\r')
							time.sleep(1)
							itemp2=itemp2-1
						if (cmd_all_output.returncode == 0):
							ibt_result=True
						else:
							ibt_result=False

						# BT藍牙喇叭連線是否成功
						if (ibt_result == True):
							idatetime=datetime.datetime.now()
							print("\033[32m[",idatetime,"] \033[37mBT(",ibt,") connection successful","                    ",sep='',end='\n')
							print("[",idatetime,"] BT (",ibt,") connection successful",sep='',end='\n',file=Log_File)
						else:
							idatetime=datetime.datetime.now()
							print("\033[32m[",idatetime,"] \033[91mBT(",ibt,") connection failed","                    ",sep='',end='\n')
							print("[",idatetime,"] BT(",ibt,") connection failed",sep='',end='\n',file=Log_File)

				else:
					idatetime=datetime.datetime.now()
					print("\033[32m[",idatetime,"] \033[91mBT(",ibt,") detection failed","                    ",sep='',end='\n')
					print("[",idatetime,"] BT(",ibt,") detection failed",sep='',end='\n',file=Log_File)

			else:
				idatetime=datetime.datetime.now()
				print("\033[32m[",idatetime,"] \033[91m","BT(",ibt,") hasn't paired with NB","                    ",sep='',end='\n')
				print("[",idatetime,"] ","BT(",ibt,") hasn't paired with NB",sep='',end='\n',file=Log_File)


    		# 如果Wi-Fi與BT都正常
			if (iap_result == True) & (ibt_result == True):

				# 執行Wi-Fi: ping -t 192.168.1.1
				idatetime=datetime.datetime.now()
				print("\033[32m[",idatetime,"] \033[37mWi-Fi: ping -t 192.168.1.1",sep='',end='\n')
				print("[",idatetime,"] Wi-Fi: ping -t 192.168.1.1",sep='',end='\n',file=Log_File)

				# 取得螢幕解析度
				screen_width, screen_height = pyautogui.size()

				#開啟新的cmd視窗
				WiFi_cmd = 'start cmd /k ping -t 192.168.1.1'
				subprocess.Popen(WiFi_cmd,shell=True)
				time.sleep(4)
				WiFi_window=getwindow.getWindowsWithTitle("cmd")[0]
				WiFi_window_x=0
				WiFi_window_y=0
				WiFi_window_width=round(screen_width/2)
				WiFi_window_height=round(screen_height)-45
				WiFi_window.moveTo(WiFi_window_x,WiFi_window_y)
				WiFi_window.resizeTo(WiFi_window_width,WiFi_window_height)


				# 執行BT: 播放音樂
				idatetime=datetime.datetime.now()
				print("\033[32m[",idatetime,"] \033[37mBT: play music.mp3",sep='',end='\n')
				print("[",idatetime,"] BT: play music.mp3",sep='',end='\n',file=Log_File)
				win32api.ShellExecute(0,'open','music.mp3','','',1)
				time.sleep(4)

				# 確認是否預設程式播放mp3, 如果是則調整視窗大小與位置
				Music_Window_Title="媒體播放器"
				hwnd=win32gui.FindWindow(None,Music_Window_Title)
				if (hwnd != 0):
					BT_window=getwindow.getWindowsWithTitle(Music_Window_Title)[0]
					BT_window_x=round(screen_width/2)
					BT_window_y=0
					BT_window_width=round(screen_width/2)
					BT_window_height=round(screen_height)-45
					BT_window.moveTo(BT_window_x,BT_window_y)
					BT_window.resizeTo(BT_window_width,BT_window_height)
				else:
					Music_Window_Title="Media Player"
					hwnd=win32gui.FindWindow(None,Music_Window_Title)
					if (hwnd != 0):
						BT_window=getwindow.getWindowsWithTitle(Music_Window_Title)[0]
						BT_window_x=round(screen_width/2)
						BT_window_y=0
						BT_window_width=round(screen_width/2)
						BT_window_height=round(screen_height)-45
						BT_window.moveTo(BT_window_x,BT_window_y)
						BT_window.resizeTo(BT_window_width,BT_window_height)

				itemp=0
				idatetime=datetime.datetime.now()
				print("[",idatetime,"] Wi-Fi and BT are running",sep='',end='\n',file=Log_File)
				while itemp<irun:
					print("\033[32m[",datetime.datetime.now(),"] \033[37mWi-Fi and BT are running...waiting ",irun,"                    ",sep='',end='\r')
					time.sleep(1)
					irun=irun-1
				print("\033[32m[",idatetime,"] \033[37mWi-Fi and BT are running","                    ",sep='',end='\n')
				idatetime=datetime.datetime.now()
				print("\033[32m[",idatetime,"] \033[37mWi-Fi and BT are finished","                    ",sep='',end='\n')
				print("[",idatetime,"] Wi-Fi and BT are finished",sep='',end='\n',file=Log_File)

				itemp=0
				while itemp<ishutdown:
					idatetime=datetime.datetime.now()
					print("\033[32m[",idatetime,"] \033[37mNB shutdown...waiting ",ishutdown,"                    ",sep='',end='\r')
					time.sleep(1)
					ishutdown=ishutdown-1

				Config_Setting=open(Config_File, mode='r', encoding='utf-8')
				issid=Config_Setting.readline()
				ibt=Config_Setting.readline()
				istart=Config_Setting.readline()
				istart=int(istart.replace('\n', ''))
				iend=Config_Setting.readline()
				iloop1=Config_Setting.readline()
				iloop2=Config_Setting.readline()
				irun=Config_Setting.readline()
				ishutdown=Config_Setting.readline()
				Config_Setting.close()

				istart=istart+1
				istart=str(istart)

				iConfig=[issid,ibt,istart,'\n',iend,iloop1,iloop2,irun,ishutdown]
				Config_Setting=open(Config_File, mode='w', encoding='utf-8')
				Config_Setting.writelines(iConfig)
				Config_Setting.close()

				end_time=datetime.datetime.now()
				print("\033[32m[",end_time,"] \033[37m","Computer is shutdown (Spent Time:",str(end_time-start_time).split(".")[0],")","                    ",sep='',end='\n')
				print("[",end_time,"] Computer is shutdown (Spent Time:",str(end_time-start_time).split(".")[0],")",sep='',end='\n',file=Log_File)
				Log_File.close()
				os.system("shutdown -s -t 1")
			else:
				Log_File.close()

			Log_File.close()

		else:
			idatetime=datetime.datetime.now()
			print("\033[32m[",idatetime,"] \033[91m","There is no Wi-Fi interface to connect AP(",issid,")",sep='',end='\n')
			print("[",idatetime,"] ","There is no Wi-Fi interface to connect AP(",issid,")",sep='',end='\n',file=Log_File)

	else:
		os.system("cls")
		idatetime=datetime.datetime.now()
		Log_File=open(Log_File, mode='a+')
		print("\033[32m[",idatetime,"] \033[37m","Power ON-OFF test is finished (start loop > end loop)",sep='',end='\n')
		print("[",idatetime,"] Power ON-OFF test is finished (start loop > end loop)",sep='',end='\n',file=Log_File)
		Log_File.close()

	Config_Setting.close()
	Log_File.close()


#處理異常訊息
except Exception as e:
	print("\n\033[93m")
	print("Exception Type:\t\t",str(Exception))
	print("Exception Message:\t",str(e))
	print("Exception Line:\t\t",str(e.__traceback__.tb_lineno))
	print("\033[0m")
finally:
	Config_Setting.close()
	Log_File.close()
	print("\033[0m")
	os.system("pause")