import xlwings as xw 
import os

ret=os.path.exists(r"E:\preconf.sh")
if ret:
	print("file is exit , del it!\n")
	os.remove(r"E:\preconf.sh")
else:
	print("file not exit!\n")

preconf=open(r"E:\preconf.sh","a")
preconf.write("#!/bin/sh\r\n")
preconf.write("prolinecmd clear 1\n")

app=xw.App(visible=True,add_book=False)

filepath=r"E:\103_MTK预配置.xlsx"
prefile=app.books.open(filepath)

presheet=prefile.sheets("MTK")

for i in range(3,75):
	indexB="B"+str(i)
	indexC="C"+str(i)
	print("The line {0} value is {1} {2}".format(i,presheet.range(indexB).value,presheet.range(indexC).value))
	if presheet.range(indexB).value=="MAC地址":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="sys mac "+ presheet.range(indexC).value +" -n\n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="设备型号":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd productclass set \""+ presheet.range(indexC).value +"\" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="设备OUI":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd manufacturerOUI set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="设备系列号":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd serialnum set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="硬件版本号":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd hwver set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="皮肤":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd ResDir set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="运营商":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd ResModelType set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="上行方式":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd WanTransMode set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="PON SN":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd xponsn set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="普通用户":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd webAccount set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="普通密码":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd webpwd set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="管理员用户":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd cfeusrname set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="管理员密码":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd cfepwd set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口1 开关":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan1_Active set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口1 VLAN模式":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan1_VLANMode set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口1 VLANID":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan1_VLANID set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口1 802.1p":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan1_dot1pData set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口1 模式":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan1_WanMode set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口1 MTU":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan1_MTU set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口1 链接方式":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan1_LinkMode set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口1 用户名":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan1_USERNAME set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口1 密码":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan1_PASSWORD set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口1 承载业务":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan1_ServiceList set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口1 绑定端口":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan1_BindPort set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口2 标志位":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan2_Flag set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口2 开关":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan2_Active set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口2 VLANID":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan2_VLANID set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口2 802.1p":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan2_dot1pData set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口2 模式":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan2_WanMode set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口2 MTU":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan2_MTU set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口2 链接方式":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan2_LinkMode set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口2 承载业务":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan2_ServiceList set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口2 绑定端口":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan2_BindPort set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口3 标志位":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan3_Flag set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口3 开关":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan3_Active set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口3 VLANID":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan3_VLANID set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口3 802.1p":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan3_dot1pData set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口3 模式":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan3_WanMode set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口3 MTU":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan3_MTU set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口3 链接方式":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan3_LinkMode set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口3 承载业务":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan3_ServiceList set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口3 绑定端口":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan3_BindPort set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口4 标志位":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan4_Flag set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口4 开关":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan4_Active set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口4 VLANID":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan4_VLANID set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口4 802.1p":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan4_dot1pData set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口4 模式":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan4_WanMode set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口4 MTU":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan4_MTU set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口4 链接方式":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan4_LinkMode set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口4 承载业务":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan4_ServiceList set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="子接口4 绑定端口":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd Wan4_BindPort set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="ALG功能":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd ALGSwitch set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="TR069 URL":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd acsUrl set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="ACS用户名":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd acsUserName set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="ACS密码":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd acsPassword set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="CPE用户名":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd conReqUserName set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="CPE密码":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd conReqPassword set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="语音工作模式":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd SIPProtocol set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="SIP域":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd SIPDomain set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="注册服务器":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd RegServer set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="代理服务器":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd IPProxyAddr set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="出局代理服务器":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd SIPOutroxyAddr set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="备用注册服务器":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd SBRegServer set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="备用代理服务器":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd SBSIPProxyAddr set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="备用出局代理服务器":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd SBOutProxyAddr set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="会话更新周期（秒）":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd ACCT_TIMER set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="注册周期（秒）":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd RegExpire set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="回音消除启用":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd ECEnable set "+ presheet.range(indexC).value +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="编解码优先顺序":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd priority set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="同步话机时间":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd SyncCaller set "+ str(int(presheet.range(indexC).value)) +" \n"
			preconf.write(cmd)
			print(cmd)
	elif presheet.range(indexB).value=="基本数图表":
		if presheet.range(indexC).value==None:
			continue;
		else:
			cmd="prolinecmd DigitMap1 set \""+ presheet.range(indexC).value +"\" \n"
			preconf.write(cmd)
			print(cmd)

preconf.write("prolinecmd restore default\n")

prefile.close()
app.quit()
preconf.close()