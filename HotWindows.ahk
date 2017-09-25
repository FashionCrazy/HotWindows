DetectHiddenWindows,On
#WinActivateForce
#Persistent
#SingleInstance force
#UseHook
#InstallKeybdHook
#ErrorStdOut
ComObjError(false)
SetBatchLines -1

if not A_IsUnicode
{
	MsgBox,编译出错请使用AutoHotkeyU32运行
	ExitApp
}

;<<<<<<<<<<<<默认值>>>>>>>>>>>>
WS_EX_APPWINDOW = 0x40000 ; provides a taskbar button
WS_EX_TOOLWINDOW = 0x80 ; removes the window from the alt-tab list
GW_OWNER = 4
Path_data=%A_ScriptDir%\HotWindows.mdb	;数据库地址

;<<<<<<<<<<<<WIN10 WIN8中重要的设置值>>>>>>>>>>>>
;SplashImage,F:\Git\HotWindows\alipayhotwin12.png,b x0 y0 ; ,下面文本,上面的文本,窗口标题
Progress,,初始化,初始化请稍等...,HotWindows
;<<<<<<<<<<<<预设与配置>>>>>>>>>>>>
Menu,Tray,NoStandard
Show_modes=TrayTip,ListView
Hot_keys=Space,Tab
Menu,Show_mode,Add,Traytip,Traytip
Menu,Show_mode,Add,listview,listview
loop,Parse,Hot_keys,`,
	Menu,Hot_key,Add,%A_LoopField%,Hot_key
Menu,Tray,Add,开机启动,Auto
Menu,Tray,Add,输入保护,Boot
Menu,Tray,Add,支持作者,Support
Menu,Tray,Add
Menu,Tray,Add,激活热键,:Hot_key
Menu,Tray,Add,显示方式,:Show_mode
Menu,Dele_mdb,Add,清除窗口记录,Dele_mdb_Gui
Menu,Dele_mdb,Add,清除程序记录,Dele_mdb_Exe
Menu,Dele_mdb,Add,清除所有记录,Dele_mdb
Menu,Tray,Add,清除记录,:Dele_mdb
Menu,Tray,Add,添加程序,Add_exe
Menu,Tray,Add
Menu,Tray,Add,重启脚本,Reload
Menu,Tray,Add,退出脚本,ExitApp
IfExist,MenuIco.icl
{
	Menu,Tray,Icon,激活热键,MenuIco.icl,7
	Menu,Tray,Icon,支持作者,MenuIco.icl,4
	Menu,Tray,Icon,显示方式,MenuIco.icl,1
	Menu,Tray,Icon,添加程序,MenuIco.icl,8
	Menu,Tray,Icon,清除记录,MenuIco.icl,2
	Menu,Tray,Icon,重启脚本,MenuIco.icl,5
	Menu,Tray,Icon,退出脚本,MenuIco.icl,6
	Menu,Tray,Icon,MenuIco.icl,9
	Menu,Tray,Icon,,,1
}
RegRead,HotRun,HKEY_CURRENT_USER,Software\Microsoft\Windows\CurrentVersion\Run,HotRun
RegRead,Show_mode,HKEY_CURRENT_USER,HotWindows,HotShow_mode	;显示方式
RegRead,Hot_Set_key,HKEY_CURRENT_USER,HotWindows,HotHot_key	;激活热键
RegRead,Boot,HKEY_CURRENT_USER,HotWindows,Hotboot	;输入保护
RegRead,Path_list,HKEY_CURRENT_USER,HotWindows,Path_list	;自定义程序添加规则
if not Path_list
	RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,Path_list,%A_Desktop%\*.lnk`n
RegRead,Path_list,HKEY_CURRENT_USER,HotWindows,Path_list	;自定义程序添加规则
if not Hot_Set_key{
	Hot_Set_key=Space
	RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,HotHot_key,% Hot_Set_key
	Menu,Hot_key,ToggleCheck,Space
}else{
	Menu,Hot_key,ToggleCheck,%Hot_Set_key%
}
if not Show_mode{
	Show_mode=ListView
	RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,HotShow_mode,% Show_mode
	Menu,Show_mode,ToggleCheck,listview
}else{
	Menu,Show_mode,ToggleCheck,%Show_mode%
}
IfExist,%HotRun%
{
	RegWrite,REG_SZ,HKEY_CURRENT_USER,Software\Microsoft\Windows\CurrentVersion\Run,HotRun,%A_ScriptFullPath%
	Menu,Tray,ToggleCheck,开机启动
}
if not Boot{
	Boot:=1
	RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,Hotboot,% Boot
	Menu,Tray,ToggleCheck,输入保护
}else if (Boot="1"){
	Menu,Tray,ToggleCheck,输入保护
}

;<<<<<<<<<<<<热键创建>>>>>>>>>>>>
Layout=qwertyuiopasdfghjklzxcvbnm
loop,Parse,Layout
{
	Layouts:=A_LoopField
	loop,Parse,Hot_keys,`,
	{
		Hotkey,~%A_LoopField% & %Layouts%,Layout
		if (Hot_Set_key!=A_LoopField)
			Hotkey,~%A_LoopField% & %Layouts%,off
	}
}
SysGet,Width,16
SysGet,Height,17
ListWidth:=Width/4

;<<<<<<<<<<<<声明全局变量>>>>>>>>>>>>
global Path_data,Show_mode,K_ThisHotkey,WHERE_list,Path_list,Ger,Gers,WS_EX_TOOLWINDOW,WS_EX_APPWINDOW,GW_OWNER,complete

;<<<<<<<<<<<<检查更新>>>>>>>>>>>>
UpdateInfo:=Git_Update("https://github.com/liumenggit/HotWindows","Show")


;<<<<<<<<<<<<GUI>>>>>>>>>>>>
;http://new.cnzz.com/v1/login.php?siteid=1261658612
Gui,+AlwaysOnTop +Border -SysMenu +ToolWindow +LastFound +HwndMyGuiHwnd
Gui,Add,ListView,w%ListWidth% r9 xm ym AltSubmit gHot_ListView,编号|标题	;
Gui,Add,StatusBar
WinSet,Transparent,200,ahk_id %MyGuiHwnd%

;<<<<<<<<<<<<创建SQL表>>>>>>>>>>>>
IfNotExist,%Path_data%
{
	Catalog:=ComObjCreate("ADOX.Catalog")
	Catalog.Create("Provider='Microsoft.Jet.OLEDB.4.0';Data Source=" Path_data)
	SQL_Run("CREATE TABLE Activate(Title varchar(255),pinyin varchar(255),Times varchar(255),Add_Time varchar(255))")	;添加程序数据库表
	SQL_Run("CREATE TABLE Nowlist(Title varchar(255),Pid varchar(255),Path varchar(255))")	;添加程序数据库表
	SQL_Run("CREATE TABLE Quick(Title varchar(255),Pid varchar(255),Path varchar(255))")	;添加程序数据库表
}else{
	SQL_Run("Delete FROM Nowlist")
}

;<<<<<<<<<<<<加载列表>>>>>>>>>>>>
Load_list()	;创建初始程列表
TrayTip,HotWindows,% "准备完成开始使用`n当前版本号：" UpdateInfo.Edition "`n捐赠支付宝：rrsyycm@163.com",,1
Menu,Tray,Tip,% "HotWindows`n版本:" UpdateInfo.Edition
;<<<<<<<<<<<<主要循环>>>>>>>>>>>>
loop{
	WinGet,Wina_ID,ID,A
	Load_exe(Wina_ID)
	WinWaitNotActive,ahk_id %Wina_id%
	Load_exe(Wina_ID)
}
return

;<<<<<<<<<<<<主要功能的标签>>>>>>>>>>>>
Layout:
Critical
if complete{
	Critical off
	return
}
StringRight,H_ThisHotkey,A_ThisHotkey,1
K_ThisHotkey:=K_ThisHotkey H_ThisHotkey
StrLens := StrLen(K_ThisHotkey)
if (StrLens="1"){
	loop,9{
		Hotkey,%A_Index%,Table
		Hotkey,%A_Index%,On
	}
	SetTimer,Key_wait,1
}
SQL_List("SELECT Activate.title,Activate.times,t1.pid,t1.path FROM Activate LEFT JOIN (SELECT * FROM Nowlist UNION SELECT * FROM Quick) AS t1 ON Activate.title = t1.title WHERE (t1.pid IS NOT NULL OR t1.path IS NOT NULL) and Activate.pinyin LIKE '%" K_ThisHotkey "%' ORDER BY Activate.Times +- 1 DESC,t1.pid DESC")
if WHERE_list.Length() and K_ThisHotkey{
	Show_list(WHERE_list)
	if (WHERE_list.Length()="1"){
		Critical off
		Activate("1")
		Send {%Hot_Set_key% Up}
	}
}else{
	Critical off
	Cancel()
}
return


Table:
Activate(A_ThisHotkey)
return

Key_wait:
SetTimer,Key_wait,off
KeyWait,%Hot_Set_key%,L
if not K_ThisHotkey{
	Cancel()
	return
}
if (EventInfo<>"0") and (Show_mode="ListView"){
	if (Boot="1") and (StrLens="1")
		Cancel()
	if (Boot="2")
		Activate(EventInfo)
	if (Boot="1") and (StrLens>"1")
		Activate(EventInfo)
	return
}
if (Boot="1") and (StrLens="1")	;开启了输入保护什么也没有发生
	Cancel()
if (Boot="2")	;没有开启输入保护激活第一个
	Activate("1")
if (Boot="1") and (StrLens>"1")	;开启了输入保护发生了事情
	Activate("1")
Cancel()
return

Add_exe:
	Gui,New
	Gui,Add_exe:New
	Gui,Add_exe:+LabelMyAdd +AlwaysOnTop
	Gui,Add_exe:Add,Text,xm,添加程序请将文件拖入本窗口
	Gui,Add_exe:Add,ListView,xm w%ListWidth% vAdd_list r9,名称|路径
	Gui,Add_exe:Add,Text,xm,此处添加程序目录c:\Users\*.exe或c:\Users\*.lnk
	Gui,Add_exe:Add,Edit,xm w%ListWidth% r5 vPath_list,%Path_list%
	Gui,Add_exe:Add,Button,xm Section gDele_exe,删除选择程序(&D)
	Gui,Add_exe:Add,Button,ys gSubmit_exe,保存规则(&S)
	Gui,Add_exe:Show,,添加程序到热启动列表
	Add_list()
return

Submit_exe:
	TrayTip,HotWindows,等待操作完成,,1
	Gui,Submit,NoHide
	RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,Path_list,%Path_list%
	Load_list()
	Add_list()
	TrayTip,HotWindows,保存完成,,1
return

Dele_exe:
	Gui,ListView,Add_list
	RowNumber=0
	loop
	{
		RowNumber:=LV_GetNext(RowNumber)
		if not RowNumber
			break
		LV_GetText(dPath,RowNumber,2)
		SQL_Run("Delete FROM Quick WHERE Path='" dPath "'")
	}
	Add_list()
	TrayTip,HotWindows,删除完成,,3
	Gui,ListView,Hot_ListView
return

MyAddDropFiles:
	loop,Parse,A_GuiEvent,`n
		Add_quick(A_LoopField)
	Add_list()
	TrayTip,HotWindows,添加完成,,1
return

Hot_ListView:
	EventInfo:=LV_GetNext()
return
;<<<<<<<<<<<<窗口函数>>>>>>>>>>>>
Add_list(){
	Recordset := ComObjCreate("ADODB.Recordset")
	Recordset.Open("SELECT Title,Path FROM Quick","Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . Path_data . "")
	GuiControl,-Redraw,Add_list
	LV_Delete()
	while !Recordset.EOF
	{
		IfExist,% Recordset.Fields["Path"].Value
			LV_Add("" ,Recordset.Fields["Title"].Value,Recordset.Fields["Path"].Value)
		else
			SQL_Run("Delete FROM Quick WHERE Path='" Recordset.Fields["Path"].Value "'")
		Recordset.MoveNext()
	}
	LV_ModifyCol()
	GuiControl,+Redraw,Add_list
}

Activate(WHERE_time){
	Cancel()
	Activate:=WHERE_list[WHERE_time].PID
	Title:=WHERE_list[WHERE_time].Title
	Path:=WHERE_list[WHERE_time].Path
	complete:=1
	if Activate{
		WinActivate,ahk_id %Activate%
	}else{
		Try{
			RunWait %Path%
	}catch e{
	KeyWait,Space
	complete:=
	return
}
}
if not SQL_Get("SELECT Times FROM Activate WHERE Title='" Title "'")
	SQL_Run("Insert INTO Activate (Title,pinyin,Times,Add_Time) VALUES ('" Title "','" zh2py(Title) "','1','" A_Now "')")
else
	SQL_Run("UPDATE Activate SET Times = Times+1 , Add_Time = '" A_Now "' WHERE Title='" Title "'")
KeyWait,Space
complete:=
}

Cancel(){
	if K_ThisHotkey
		loop,9
			Hotkey,%A_Index%,off
	if Show_mode=ListView
		Gui,Cancel
	else
		TrayTip
	K_ThisHotkey:=
	Send {%Hot_Set_key% Up}
}

Show_list(WHERE_list){
	if Show_mode=ListView
	{
		GuiControl,-Redraw,MyListView
		LV_Delete()
		ImageListID:=IL_Create(WHERE_list.Length())
		LV_SetImageList(ImageListID)
		For k,v in WHERE_list
		{
			if v.Pid
				Level=○
			else
				Level=●
			LV_Add("Icon" . IL_Add(ImageListID,v.Path,1),k Level,v.Title)
		}
		LV_ModifyCol()
		SB_SetText("按键历史：" . K_ThisHotkey . "")
		LV_Modify(1,"Select")
		LV_Modify(1,"Focus")
		GuiControl,+Redraw,MyListView
		Gui,Show,AutoSize CEnter,HotWindows
	}else{
		Tip_list:=K_ThisHotkey
		For k,v in WHERE_list
		{
			if v.Pid
				Level=○
			else
				Level=●
			Tip_list:=Tip_list "`n" k Level SubStr(v.Title,"1","25")
		}
		TrayTip,,%Tip_list%
	}
	if not K_ThisHotkey
		Cancel()
}

;<<<<<<<<<<<<生成数据>>>>>>>>>>>>
Load_list(){
	Suspend,On
	windowList =
	DetectHiddenWindows, Off ; makes DllCall("IsWindowVisible") unnecessary
	WinGet, windowList, List ; gather a list of Running programs
	loop, %windowList%
	{
		Load_exe(windowList%A_Index%)
		Progress,% Gers:=90//windowList*A_Index ,% Gers "%",构建当前窗口信息...,HotWindows
	}
	;从数据库删除不存在的程序
	Recordset := ComObjCreate("ADODB.Recordset")
	Recordset.CursorLocation:="3"
	Recordset.Open("SELECT Path FROM Quick","Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . Path_data . "")
	while !Recordset.EOF
	{
		Progress,% Gers:=90+10//Recordset.RecordCount*A_Index ,% Gers "%",删除不存在程序...,HotWindows
		Quick_Path:=Recordset.Fields["Path"].Value
		IfNotExist,%Quick_Path%
			SQL_Run("Delete FROM Quick WHERE Path='" Quick_Path "'")
		Recordset.MoveNext()
	}
	;删除两天前的沉淀数据
	SQL_Run("Delete FROM Activate WHERE Add_Time < '" A_Now - 172800 "'")
	;根据配置添加
	loop,Parse,Path_list,`n
		loop,%A_LoopField%
			Add_quick(A_LoopFileLongPath)
	Add_list()
	Progress,100
	Progress,Off
	;SplashImage,Off
	Suspend,Off
}
Load_exe(windowID){
	AltTabTotalNum := 0 ; the number of windows found
	AltTabListID_1 =    ; hwnd from last active windows
	AltTabListID_2 =    ; hwnd from previous active windows
	ownerID := windowID
	loop {
		ownerID := Decimal_to_Hex( DllCall("GetWindow", "UInt", ownerID, "UInt", GW_OWNER))
	} Until !Decimal_to_Hex( DllCall("GetWindow", "UInt", ownerID, "UInt", GW_OWNER))
	ownerID := ownerID ? ownerID : windowID
	if (Decimal_to_Hex(DllCall("GetLastActivePopup", "UInt", ownerID)) = windowID)
	{
		WinGet, es, ExStyle, ahk_id %windowID%
		if (!((es & WS_EX_TOOLWINDOW) && !(es & WS_EX_APPWINDOW)) && !IsInvisibleWin10BackgroundAppWindow(windowID))
		{
			AltTabTotalNum ++
			AltTabListID_%AltTabTotalNum% := windowID
			WinGet,Path,ProcessPath,ahk_id %windowID%
			WinGetTitle,Title,ahk_id %windowID%
			SplitPath,Path,OutFileName,OutDir,OutExtension,OutNameNoExt,OutDrive
			;IfNotInString,Path,%A_WinDir%
			SQL_Run("Delete FROM Nowlist WHERE Path LIKE '%" OutFileName "'")
			;MsgBox % Title
			if SQL_Get("SELECT COUNT(*) FROM Nowlist WHERE Title='" Title "'")
				return
			Add_quick(Path)
			SQL_Run("Insert INTO Nowlist (Title,PID,Path) VALUES ('" Title "','" windowID "','" Path "')")
			if not SQL_Get("SELECT Times FROM Activate WHERE Title='" Title "'")
				SQL_Run("Insert INTO Activate (Title,pinyin,Times,Add_Time) VALUES ('" Title "','" zh2py(Title) "','1','" A_Now "')")
		}
	}
}

Add_quick(Path){
	IfNotInString,Path,%A_WinDir%
	{
		SplitPath,Path,OutFileName,OutDir,OutExtension,OutNameNoExt,OutDrive
		if not SQL_Get("SELECT COUNT(*) FROM Quick WHERE Path='" Path "'"){
			SQL_Run("Delete FROM Quick WHERE Path='" Path "'")
			SQL_Run("Insert INTO Quick (Title,Path) VALUES ('" OutNameNoExt "','" Path "')")
		}
		if not SQL_Get("SELECT Times FROM Activate WHERE Title='" OutNameNoExt "'")
			SQL_Run("Insert INTO Activate (Title,pinyin,Times,Add_Time) VALUES ('" OutNameNoExt "','" zh2py(OutNameNoExt) "','1','" A_Now "')")
	}
}

Decimal_to_Hex(var) {
	SetFormat, IntegerFast, H
	var += 0
	var .= ""
	SetFormat, Integer, D
	return var
}
IsInvisibleWin10BackgroundAppWindow(hWindow) {
	result := 0
	VarSetCapacity(cloakedVal, A_PtrSize) ; DWMWA_CLOAKED := 14
	hr := DllCall("DwmApi\DwmGetWindowAttribute", "Ptr", hWindow, "UInt", 14, "Ptr", &cloakedVal, "UInt", A_PtrSize)
	if !hr ; returns S_OK (which is zero) on success. Otherwise, it returns an HRESULT error code
		result := NumGet(cloakedVal) ; omitting the "&" performs better
	return result ? true : false
}

GetIconCount(file){
	Menu, test, add, test, handle
	loop
	{
		try {
			id++
		Menu, test, Icon, test, % file, % id
	} catch error {
	break
}
}
return id-1
}
handle:
return
;<<<<<<<<<<<<MENU的功能>>>>>>>>>>>>
Traytip:
	if (Show_mode = "TrayTip")
		return
	RegRead,Bubble,HKEY_CURRENT_USER,SOFTWARE\Policies\Microsoft\Windows\Explorer,EnableLegacyBalloonNotifications
	if not Bubble {
		MsgBox,4,重要设置,点击'Yes'资源管理器会重启如有重要进程请保存后再次设置。
		IfMsgBox Yes
		{
			RegWrite,REG_DWORD,HKEY_CURRENT_USER,SOFTWARE\Policies\Microsoft\Windows\Explorer,EnableLegacyBalloonNotifications,1
			RunWait %comspec% /c "taskkill /f /im explorer.exe",,Hide
			Run %comspec% /c "start c:\Windows\explorer.exe",,Hide
		}else{
			return
		}
	}
	Menu,Show_mode,ToggleCheck,Traytip
	Menu,Show_mode,Uncheck,listview
	Show_mode:="TrayTip"
	RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,HotShow_mode,TrayTip
return

Support:
	Run "https://github.com/liumenggit/HotWindows#捐赠开发者"
return

Auto:
	IfExist,%HotRun%
		RegDelete,HKEY_CURRENT_USER,Software\Microsoft\Windows\CurrentVersion\Run,HotRun
	else
		RegWrite,REG_SZ,HKEY_CURRENT_USER,Software\Microsoft\Windows\CurrentVersion\Run,HotRun,%A_ScriptFullPath%
	Menu,Tray,ToggleCheck,开机启动
return

listview:
	if (Show_mode = "Listview")
		return
	Show_mode:="Listview"
	Menu,Show_mode,Uncheck,Traytip
	Menu,Show_mode,ToggleCheck,listview
	RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,HotShow_mode,Listview
return
Hot_key:
	loop,Parse,Layout
	{
		Hotkey,~%Hot_Set_key% & %A_LoopField%,off
		Hotkey,~%A_ThisMenuItem% & %A_LoopField%,On
	}
	Hot_Set_key:=A_ThisMenuItem
	loop,Parse,Hot_keys,`,
		Menu,Hot_key,Uncheck,%A_LoopField%
	Menu,Hot_key,ToggleCheck,%Hot_Set_key%
	RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,HotHot_key,%A_ThisMenuItem%
return

Boot:
	RegRead,Boot,HKEY_CURRENT_USER,HotWindows,Hotboot	;输入保护
	if Boot=1
		Boot=2
	else
		Boot=1
	RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,Hotboot,%Boot%
	Menu,Tray,ToggleCheck,输入保护
return

Dele_mdb:
	TrayTip,HotWindows,等待操作完成,,1
	SQL_Run("Delete FROM Activate")
	SQL_Run("Delete FROM Nowlist")
	SQL_Run("Delete FROM Quick")
	RegDelete,HKEY_CURRENT_USER,HotWindows,HotStyles
	RegDelete,HKEY_CURRENT_USER,HotWindows,Path_list
	Load_list()
	TrayTip,HotWindows,已经清除所有记录,,1
return

Dele_mdb_Gui:
	TrayTip,HotWindows,等待操作完成,,1
	SQL_Run("Delete FROM Activate")
	Load_list()
	TrayTip,HotWindows,已经清除窗口记录,,1
return

Dele_mdb_Exe:
	TrayTip,HotWindows,等待操作完成,,1
	SQL_Run("Delete FROM Quick")
	RegDelete,HKEY_CURRENT_USER,HotWindows,Path_list
	Load_list()
	TrayTip,HotWindows,已经清除程序记录,,1
return


Reload:
	Reload
ExitApp:
	ExitApp

	;<<<<<<<<<<<<SQL函数>>>>>>>>>>>>
	SQL_List(SQL){
	Recordset := ComObjCreate("ADODB.Recordset")
	Recordset.Open(SQL,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . Path_data . "")
	WHERE_list := Object()
	WHERE_time :=
	while !Recordset.EOF
	{
		wPID:=Recordset.Fields["PID"].Value
		wPath:=Recordset.Fields["Path"].Value
		if wPID
		{
			IfWinNotExist,ahk_id %wPID%
			{
				SQL_Run("Delete FROM Nowlist WHERE PID='" Recordset.Fields["PID"].Value "'")
				Recordset.MoveNext()
				continue
			}
		}
		else if wPath
		{
			IfNotExist,%wPath%
			{
				SQL_Run("Delete FROM Nowlist WHERE Path='" Recordset.Fields["Path"].Value "'")
				Recordset.MoveNext()
				continue
			}
		}
		;MsgBox % Recordset.Fields["Title"]
		WHERE_time++
		WHERE_list[WHERE_time]:={Title:Recordset.Fields["Title"].Value,PID:Recordset.Fields["PID"].Value,Path:Recordset.Fields["Path"].Value,GetStyle:Recordset.Fields["GetStyle"].Value,Times:Recordset.Fields["Times"].Value}
		Recordset.MoveNext()
	}
return
}
SQL_Run(SQL){	;向数据库运行命令
	Recordset := ComObjCreate("ADODB.Recordset")
	Try Recordset.Open(SQL,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . Path_data . "")
		catch e
			return
}
SQL_Get(SQL){	;向数据库运行命令请求返回
	Recordset := ComObjCreate("ADODB.Recordset")
	Try Recordset.Open(SQL,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . Path_data . "")
		catch e
			return 0
	Try return Recordset.Fields[0].Value
		catch e
			return 0
}

zh2py(str)
{
	static FirstTable  := [ 0xB0C5, 0xB2C1, 0xB4EE, 0xB6EA, 0xB7A2, 0xB8C1, 0xB9FE, 0xBBF7, 0xBFA6, 0xC0AC, 0xC2E8
		, 0xC4C3, 0xC5B6, 0xC5BE, 0xC6DA, 0xC8BB, 0xC8F6, 0xCBFA, 0xCDDA, 0xCEF4, 0xD1B9, 0xD4D1, 0xD7FA ]
	static FirstLetter := StrSplit("ABCDEFGHJKLMNOPQRSTWXYZ")
	static SecondTable := [ StrSplit("CJWGNSPGCGNEGYPBTYYZDXYKYGTZJNMJQMBSGZSCYJSYYFPGKBZGYDYWJKGKLJSWKPJQHYJWRDZLSYMRYPYWWCCKZNKYYG")
		, StrSplit("TTNGJEYKKZYTCJNMCYLQLYPYSFQRPZSLWBTGKJFYXJWZLTBNCXJJJJTXDTTSQZYCDXXHGCKBPHFFSSTYBGMXLPBYLLBHLX")
		, StrSplit("SMZMYJHSOJNGHDZQYKLGJHSGQZHXQGKXZZWYSCSCJXYEYXADZPMDSSMZJZQJYZCJJFWQJBDZBXGZNZCPWHWXHQKMWFBPBY")
		, StrSplit("DTJZZKXHYLYGXFPTYJYYZPSZLFCHMQSHGMXXSXJYQDCSBBQBEFSJYHWWGZKPYLQBGLDLCDTNMAYDDKSSNGYCSGXLYZAYPN")
		, StrSplit("PTSDKDYLHGYMYLCXPYCJNDQJWXQXFYYFJLEJPZRXCCQWQQSBZKYMGPLBMJRQCFLNYMYQMSQYRBCJTHZTQFRXQHXMQJCJLY")
		, StrSplit("QGJMSHZKBSWYEMYLTXFSYDXWLYCJQXSJNQBSCTYHBFTDCYZDJWYGHQFRXWCKQKXEBPTLPXJZSRMEBWHJLBJSLYYSMDXLCL")
		, StrSplit("QKXLHXJRZJMFQHXHWYWSBHTRXXGLHQHFNMGYKLDYXZPYLGGSMTCFBAJJZYLJTYANJGBJPLQGSZYQYAXBKYSECJSZNSLYZH")
		, StrSplit("ZXLZCGHPXZHZNYTDSBCJKDLZAYFFYDLEBBGQYZKXGLDNDNYSKJSHDLYXBCGHXYPKDJMMZNGMMCLGWZSZXZJFZNMLZZTHCS")
		, StrSplit("YDBDLLSCDDNLKJYKJSYCJLKWHQASDKNHCSGAGHDAASHTCPLCPQYBSZMPJLPCJOQLCDHJJYSPRCHNWJNLHLYYQYYWZPTCZG")
		, StrSplit("WWMZFFJQQQQYXACLBHKDJXDGMMYDJXZLLSYGXGKJRYWZWYCLZMSSJZLDBYDCFCXYHLXCHYZJQSQQAGMNYXPFRKSSBJLYXY")
		, StrSplit("SYGLNSCMHCWWMNZJJLXXHCHSYZSTTXRYCYXBYHCSMXJSZNPWGPXXTAYBGAJCXLYXDCCWZOCWKCCSBNHCPDYZNFCYYTYCKX")
		, StrSplit("KYBSQKKYTQQXFCMCHCYKELZQBSQYJQCCLMTHSYWHMKTLKJLYCXWHEQQHTQKZPQSQSCFYMMDMGBWHWLGSLLYSDLMLXPTHMJ")
		, StrSplit("HWLJZYHZJXKTXJLHXRSWLWZJCBXMHZQXSDZPSGFCSGLSXYMJSHXPJXWMYQKSMYPLRTHBXFTPMHYXLCHLHLZYLXGSSSSTCL")
		, StrSplit("SLDCLRPBHZHXYYFHBMGDMYCNQQWLQHJJCYWJZYEJJDHPBLQXTQKWHLCHQXAGTLXLJXMSLJHTZKZJECXJCJNMFBYCSFYWYB")
		, StrSplit("JZGNYSDZSQYRSLJPCLPWXSDWEJBJCBCNAYTWGMPAPCLYQPCLZXSBNMSGGFNZJJBZSFZYNTXHPLQKZCZWALSBCZJXSYZGWK")
		, StrSplit("YPSGXFZFCDKHJGXTLQFSGDSLQWZKXTMHSBGZMJZRGLYJBPMLMSXLZJQQHZYJCZYDJWFMJKLDDPMJEGXYHYLXHLQYQHKYCW")
		, StrSplit("CJMYYXNATJHYCCXZPCQLBZWWYTWBQCMLPMYRJCCCXFPZNZZLJPLXXYZTZLGDLTCKLYRZZGQTTJHHHJLJAXFGFJZSLCFDQZ")
		, StrSplit("LCLGJDJZSNZLLJPJQDCCLCJXMYZFTSXGCGSBRZXJQQCTZHGYQTJQQLZXJYLYLBCYAMCSTYLPDJBYREGKLZYZHLYSZQLZNW")
		, StrSplit("CZCLLWJQJJJKDGJZOLBBZPPGLGHTGZXYGHZMYCNQSYCYHBHGXKAMTXYXNBSKYZZGJZLQJTFCJXDYGJQJJPMGWGJJJPKQSB")
		, StrSplit("GBMMCJSSCLPQPDXCDYYKYPCJDDYYGYWRHJRTGZNYQLDKLJSZZGZQZJGDYKSHPZMTLCPWNJYFYZDJCNMWESCYGLBTZZGMSS")
		, StrSplit("LLYXYSXXBSJSBBSGGHFJLYPMZJNLYYWDQSHZXTYYWHMCYHYWDBXBTLMSYYYFSXJCBDXXLHJHFSSXZQHFZMZCZTQCXZXRTT")
		, StrSplit("DJHNRYZQQMTQDMMGNYDXMJGDXCDYZBFFALLZTDLTFXMXQZDNGWQDBDCZJDXBZGSQQDDJCMBKZFFXMKDMDSYYSZCMLJDSYN")
		, StrSplit("SPRSKMKMPCKLGTBQTFZSWTFGGLYPLLJZHGJJGYPZLTCSMCNBTJBQFKDHBYZGKPBBYMTDSSXTBNPDKLEYCJNYCDYKZTDHQH")
		, StrSplit("SYZSCTARLLTKZLGECLLKJLQJAQNBDKKGHPJTZQKSECSHALQFMMGJNLYJBBTMLYZXDXJPLDLPCQDHZYCBZSCZBZMSLJFLKR")
		, StrSplit("ZJSNFRGJHXPDHYJYBZGDLQCSEZGXLBLGYXTWMABCHECMWYJYZLLJJYHLGNDJLSLYGKDZPZXJYYZLWCXSZFGWYYDLYHCLJS")
		, StrSplit("CMBJHBLYZLYCBLYDPDQYSXQZBYTDKYXJYYCNRJMPDJGKLCLJBCTBJDDBBLBLCZQRPYXJCJLZCSHLTOLJNMDDDLNGKATHQH")
		, StrSplit("JHYKHEZNMSHRPHQQJCHGMFPRXHJGDYCHGHLYRZQLCYQJNZSQTKQJYMSZSWLCFQQQXYFGGYPTQWLMCRNFKKFSYYLQBMQAMM")
		, StrSplit("MYXCTPSHCPTXXZZSMPHPSHMCLMLDQFYQXSZYJDJJZZHQPDSZGLSTJBCKBXYQZJSGPSXQZQZRQTBDKYXZKHHGFLBCSMDLDG")
		, StrSplit("DZDBLZYYCXNNCSYBZBFGLZZXSWMSCCMQNJQSBDQSJTXXMBLTXZCLZSHZCXRQJGJYLXZFJPHYMZQQYDFQJJLZZNZJCDGZYG")
		, StrSplit("CTXMZYSCTLKPHTXHTLBJXJLXSCDQXCBBTJFQZFSLTJBTKQBXXJJLJCHCZDBZJDCZJDCPRNPQCJPFCZLCLZXZDMXMPHJSGZ")
		, StrSplit("GSZZQLYLWTJPFSYASMCJBTZYYCWMYTZSJJLJCQLWZMALBXYFBPNLSFHTGJWEJJXXGLLJSTGSHJQLZFKCGNNNSZFDEQFHBS")
		, StrSplit("AQTGYLBXMMYGSZLDYDQMJJRGBJTKGDHGKBLQKBDMBYLXWCXYTTYBKMRTJZXQJBHLMHMJJZMQASLDCYXYQDLQCAFYWYXQHZ") ]


	static nothing := VarSetCapacity(var, 2)
	if !RegExMatch(str, "[^\x{00}-\x{ff}]")
		return str

	loop, Parse, str
	{
		StrPut(A_LoopField, &var, "CP936")
		H := NumGet(var, 0, "UChar")
		L := NumGet(var, 1, "UChar")
		if (H < 0xB0 || L < 0xA1 || H > 0xF7 || L = 0xFF)
		{
			newStr .= A_LoopField
			continue
		}

		if (H < 0xD8)//(H >= 0xB0 && H <=0xD7)
		{
			W := (H << 8) | L
			For key, value in FirstTable
			{
				if (W < value)
				{
					newStr .= FirstLetter[key]
					break
				}
			}
		}
		else
			newStr .= SecondTable[ H - 0xD8 + 1 ][ L - 0xA1 + 1 ]
	}

	return newStr
}

;作者：请勿打扰
;功能：适用GItHub项目更新，加入脚本中即可使用自己的GitHub项目地址
;介绍：根据GitHub中Commitkey获取是否更新
;注意：能够使用GitHub的朋友应该对代码都非常熟悉那么有其他需要请自行修改

Git_Update(GitUrl,GressSet:="Hide"){
	if not W_InternetCheckConnection(GitUrl)
		return
	SplitPath,GitUrl,Project_Name
	RegRead,Reg_Commitkey,HKEY_CURRENT_USER,%Project_Name%,Commitkey
	if GressSet=Show
		Progress,100,% Reg_Commitkey " >>> " Git_CcommitKey.Edition,检查更新请稍等...,% Project_Name
	Git_CcommitKey:=Git_CcommitKey(GitUrl)
	if not Git_CcommitKey.Edition{	;获取更新失败返回
		Progress,Off
		return
	}
	if not Reg_Commitkey or (Reg_Commitkey<>Git_CcommitKey.Edition){	;存在更新开始更新
		Progress,1 T Cx0 FM10,初始化下载,% Reg_Commitkey " >>> " Git_CcommitKey.Edition " 简介：" Git_CcommitKey.Commit,% Project_Name
		Git_Downloand(Git_CcommitKey,Project_Name)
	}else{
		Progress,,,暂无更新,% Project_Name
	}
	;Progress,Off
	return Git_CcommitKey
}

Git_Downloand(DownloandInfo,Project_Name){
	DownUrl:="https://github.com" DownloandInfo.Down
	SplitPath,A_ScriptName,,,,A_name
	SplitPath,DownUrl,DownName,,,OutNameNoExt
	;if not Z_Down(DownUrl,"",A_name,A_Temp "\" DownName){
	if not DownloadFile(DownUrl,A_Temp "\" DownName){
		Progress,Off
		return
	}
	UncoilUrl:=A_Temp "\" A_NowUTC
	SmartZip(A_Temp "\" DownName,UncoilUrl)
	FileDelete,% A_Temp "\" DownName
	Git_Bat(UncoilUrl "\" Project_Name "-" OutNameNoExt,Project_Name,DownloandInfo.Edition)
	ExitApp
}

Git_Bat(File,RegAdd_name,Add_Edition){
	bat=
	(LTrim
:start
	ping 127.0.0.1 -n 2>nul
	del `%1
	if exist `%1 goto start
	xcopy %File% %A_ScriptDir% /s/e/y
	reg add HKEY_CURRENT_USER\%RegAdd_name% /v Commitkey /t REG_SZ /d %Add_Edition% /f
	start %A_ScriptFullPath%
	del `%0
	)
	IfExist GitDelete.bat
		FileDelete GitDelete.bat
	FileAppend,%bat%,GitDelete.bat
	Run,GitDelete.bat,,Hide
	ExitApp
	}

SmartZip(s, o, t = 16)	;内置解压函数
{
	IfNotExist, %s%
		return, -1
	oShell := ComObjCreate("Shell.Application")
	if InStr(FileExist(o), "D") or (!FileExist(o) and (SubStr(s, -3) = ".zip"))
	{
		if !o
			o := A_ScriptDir
		else ifNotExist, %o%
				FileCreateDir, %o%
		loop, %o%, 1
			sObjectLongName := A_LoopFileLongPath
		oObject := oShell.NameSpace(sObjectLongName)
		loop, %s%, 1
		{
			oSource := oShell.NameSpace(A_LoopFileLongPath)
			oObject.CopyHere(oSource.Items, t)
		}
	}
}

Git_CcommitKey(Project_Url){
	whr:=ComObjCreate("WinHttp.WinHttpRequest.5.1")
	;whr.SetProxy("HTTPREQUEST_PROXYSETTING_PROXY","proxy_server:80","*.GitHub.com") ;https://msdn.microsoft.com/en-us/library/aa384059(v=VS.85).aspx
	whr.Open("GET",Project_Url,True)
	whr.SetRequestHeader("Content-Type","application/x-www-form-urlencoded")
	Try
	{
		whr.Send()
		whr.WaitForResponse()
		RegExMatch(whr.ResponseText,"`a)(?<=data-pjax>\n\s{8})\S{7}",NewEdition)
		RegExMatch(whr.ResponseText,"`a)\/.*\.zip",Downloand)
		RegExMatch(whr.ResponseText,"`a)(?<=class=""message"" data-pjax=""true"" title="").+(?="">)",Committitle)
		;MsgBox % NewEdition "`n" Downloand "`n" Committitle "`n-------------------------"
		return {Edition:NewEdition,Down:Downloand,Commit:Committitle}
	}catch e {
	return
}
}

W_InternetCheckConnection(lpszUrl){ ;检查FTP服务是否可连接
	FLAG_ICC_FORCE_CONNECTION := 0x1
	dwReserved := 0x0
	return, DllCall("Wininet.dll\InternetCheckConnection", "Ptr", &lpszUrl, "UInt", FLAG_ICC_FORCE_CONNECTION, "UInt", dwReserved, "Int")
}


DownloadFile(UrlToFile, SaveFileAs, Overwrite := True, UseProgressBar := True) {
	;Check if the file already exists and if we must not overwrite it
	if (!Overwrite && FileExist(SaveFileAs))
		return
	;Check if the user wants a progressbar
	if (UseProgressBar) {
		;Initialize the WinHttpRequest Object
		WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
		;Download the headers
		WebRequest.Open("HEAD", UrlToFile)
		WebRequest.Send()
		;Store the header which holds the file size in a variable:
		FinalSize := WebRequest.GetResponseHeader("Content-Length")
		;Create the progressbar and the timer
		SetTimer, __UpdateProgressBar, 100
	}
	;Download the file
	URLDownloadToFile, %UrlToFile%, %SaveFileAs%
	;Remove the timer and the progressbar because the download has finished
	if (UseProgressBar) {
		Progress, Off
		SetTimer, __UpdateProgressBar, Off
		return "True"
	}
	return

	;The label that updates the progressbar
__UpdateProgressBar:
	;Get the current filesize and tick
	CurrentSize := FileOpen(SaveFileAs,"r").Length ;FileGetSize wouldn't return reliable results
	CurrentSizeTick := A_TickCount
	;Calculate the downloadspeed
	Speed := Round((CurrentSize/1024-LastSize/1024)/((CurrentSizeTick-LastSizeTick)/1000)) "Kb/s"
	;Save the current filesize and tick for the next time
	LastSizeTick := CurrentSizeTick
	LastSize := FileOpen(SaveFileAs,"r").Length
	;Calculate percent done
	PercentDone := Round(CurrentSize/FinalSize*100)
	;Update the ProgressBar
	Progress, %PercentDone%, % Speed
return
}



Z_Down(url:="http://61.135.169.125/forbiddenip/forbidden.html", Proxy:="",e:="utf-8", File:="",byref buf:=""){
	if (!(File?o:=FileOpen(File, "w"):1) or !DllCall("LoadLibrary", "str", "wininet") or !(h := DllCall("wininet\InternetOpen", "str", "", "uint", Proxy?3:1, "str", Proxy, "str", "", "uint", 0)))
		return 0
	c:=s:=0
	if (f := DllCall("wininet\InternetOpenUrl", "ptr", h, "str", url, "ptr", 0, "uint", 0, "uint", 0x80003000, "ptr", 0, "ptr"))
	{
		if File or IsByRef(buf)
		{
			VarSetCapacity(buffer,1024,0),VarSetCapacity(bufferlen,4,0)
			loop, 5
				if (DllCall("wininet\HttpQueryInfo","uint",f, "uint", 22, "uint", &buffer, "uint", &bufferlen, "uint", 0) = 1)
				{
					Progress,+20
					y:= Trim(StrGet(&buffer)," `r`n"),q:=[]
					loop,Parse,y,`r`n
					(x:=InStr(A_LoopField,":"))?q[SubStr(A_LoopField, 1,x-1)]:=Trim(SubStr(A_LoopField, x+1)):q[A_LoopField]:=""
					if (e=0)
						Return q
					((i:= Round((fj:=q["Content-Length"])/1024)) < 1024) ?(fx:=1024,fz:= "/" i " K",percent:=i) : (fx:=1048576,fz:= "/" Round(i/1024, 1) " M",percent:=i/1024)
							,VarSetCapacity(Buf, fj, 0),DllCall("QueryPerformanceFrequency", "Int64*", i), DllCall("QueryPerformanceCounter", "Int64*", x)
					break
				}
			}
			Progress,100
			While (DllCall("Wininet.dll\InternetQueryDataAvailable", "Ptr", F, "UIntP", S, "UInt", 0, "Ptr", 0) && (S > 0)) {             
				fj	?(DllCall("Wininet.dll\InternetReadFile", "Ptr", F, "Ptr", &Buf + C, "UInt", S, "UIntP", R),C += R,DllCall("QueryPerformanceCounter", "Int64*", y),((t:=(y-x)/i) >=1)?(Test(e,Round(c/fx,2) fz " | " Round(((c-w)/1024)/t) "KB/秒",Round(c/fx/percent*100)),x:=y,w:=c):"")
:(VarSetCapacity(b, c+s, 0),DllCall("RtlMoveMemory", "ptr", &b, "ptr", &buf, "ptr", c),DllCall("wininet\InternetReadFile", "ptr", f, "ptr", &b+c, "uint", s, "uint*", r),VarSetCapacity(buf, c+=r, 0), DllCall("RtlMoveMemory", "ptr", &buf, "ptr", &b, "ptr", c))
			}
			(q?((fj=c)?"":q["Error"]:=c):""),(File?(o.rawWrite(buf, c), o.close()):""), DllCall("wininet\InternetCloseHandle", "ptr", f)
		}
	DllCall("wininet\InternetCloseHandle", "ptr", h)
	Return (File or IsByRef(buf)?q:StrGet(&buf, c>>(e="utf-16"||e="cp1200"), e))
}
Test(A,b,c){
	Progress,%c%,%b%
}
