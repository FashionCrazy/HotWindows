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
	MsgBox,���������ʹ��AutoHotkeyU32����
	ExitApp
}

;<<<<<<<<<<<<Ĭ��ֵ>>>>>>>>>>>>
WS_EX_APPWINDOW = 0x40000 ; provides a taskbar button
WS_EX_TOOLWINDOW = 0x80 ; removes the window from the alt-tab list
GW_OWNER = 4
Path_data=%A_ScriptDir%\HotWindows.mdb	;���ݿ��ַ

;<<<<<<<<<<<<WIN10 WIN8����Ҫ������ֵ>>>>>>>>>>>>
;SplashImage,F:\Git\HotWindows\alipayhotwin12.png,b x0 y0 ; ,�����ı�,������ı�,���ڱ���
Progress,,��ʼ��,��ʼ�����Ե�...,HotWindows
;<<<<<<<<<<<<Ԥ��������>>>>>>>>>>>>
Menu,Tray,NoStandard
Show_modes=TrayTip,ListView
Hot_keys=Space,Tab
Menu,Show_mode,Add,Traytip,Traytip
Menu,Show_mode,Add,listview,listview
loop,Parse,Hot_keys,`,
	Menu,Hot_key,Add,%A_LoopField%,Hot_key
Menu,Tray,Add,��������,Auto
Menu,Tray,Add,���뱣��,Boot
Menu,Tray,Add,֧������,Support
Menu,Tray,Add
Menu,Tray,Add,�����ȼ�,:Hot_key
Menu,Tray,Add,��ʾ��ʽ,:Show_mode
Menu,Dele_mdb,Add,������ڼ�¼,Dele_mdb_Gui
Menu,Dele_mdb,Add,��������¼,Dele_mdb_Exe
Menu,Dele_mdb,Add,������м�¼,Dele_mdb
Menu,Tray,Add,�����¼,:Dele_mdb
Menu,Tray,Add,��ӳ���,Add_exe
Menu,Tray,Add
Menu,Tray,Add,�����ű�,Reload
Menu,Tray,Add,�˳��ű�,ExitApp
IfExist,MenuIco.icl
{
	Menu,Tray,Icon,�����ȼ�,MenuIco.icl,7
	Menu,Tray,Icon,֧������,MenuIco.icl,4
	Menu,Tray,Icon,��ʾ��ʽ,MenuIco.icl,1
	Menu,Tray,Icon,��ӳ���,MenuIco.icl,8
	Menu,Tray,Icon,�����¼,MenuIco.icl,2
	Menu,Tray,Icon,�����ű�,MenuIco.icl,5
	Menu,Tray,Icon,�˳��ű�,MenuIco.icl,6
	Menu,Tray,Icon,MenuIco.icl,9
	Menu,Tray,Icon,,,1
}
RegRead,HotRun,HKEY_CURRENT_USER,Software\Microsoft\Windows\CurrentVersion\Run,HotRun
RegRead,Show_mode,HKEY_CURRENT_USER,HotWindows,HotShow_mode	;��ʾ��ʽ
RegRead,Hot_Set_key,HKEY_CURRENT_USER,HotWindows,HotHot_key	;�����ȼ�
RegRead,Boot,HKEY_CURRENT_USER,HotWindows,Hotboot	;���뱣��
RegRead,Path_list,HKEY_CURRENT_USER,HotWindows,Path_list	;�Զ��������ӹ���
if not Path_list
	RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,Path_list,%A_Desktop%\*.lnk`n
RegRead,Path_list,HKEY_CURRENT_USER,HotWindows,Path_list	;�Զ��������ӹ���
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
	Menu,Tray,ToggleCheck,��������
}
if not Boot{
	Boot:=1
	RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,Hotboot,% Boot
	Menu,Tray,ToggleCheck,���뱣��
}else if (Boot="1"){
	Menu,Tray,ToggleCheck,���뱣��
}

;<<<<<<<<<<<<�ȼ�����>>>>>>>>>>>>
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

;<<<<<<<<<<<<����ȫ�ֱ���>>>>>>>>>>>>
global Path_data,Show_mode,K_ThisHotkey,WHERE_list,Path_list,Ger,Gers,WS_EX_TOOLWINDOW,WS_EX_APPWINDOW,GW_OWNER,complete

;<<<<<<<<<<<<������>>>>>>>>>>>>
UpdateInfo:=Git_Update("https://github.com/liumenggit/HotWindows","Show")


;<<<<<<<<<<<<GUI>>>>>>>>>>>>
;http://new.cnzz.com/v1/login.php?siteid=1261658612
Gui,+AlwaysOnTop +Border -SysMenu +ToolWindow +LastFound +HwndMyGuiHwnd
Gui,Add,ListView,w%ListWidth% r9 xm ym AltSubmit gHot_ListView,���|����	;
Gui,Add,StatusBar
WinSet,Transparent,200,ahk_id %MyGuiHwnd%

;<<<<<<<<<<<<����SQL��>>>>>>>>>>>>
IfNotExist,%Path_data%
{
	Catalog:=ComObjCreate("ADOX.Catalog")
	Catalog.Create("Provider='Microsoft.Jet.OLEDB.4.0';Data Source=" Path_data)
	SQL_Run("CREATE TABLE Activate(Title varchar(255),pinyin varchar(255),Times varchar(255),Add_Time varchar(255))")	;��ӳ������ݿ��
	SQL_Run("CREATE TABLE Nowlist(Title varchar(255),Pid varchar(255),Path varchar(255))")	;��ӳ������ݿ��
	SQL_Run("CREATE TABLE Quick(Title varchar(255),Pid varchar(255),Path varchar(255))")	;��ӳ������ݿ��
}else{
	SQL_Run("Delete FROM Nowlist")
}

;<<<<<<<<<<<<�����б�>>>>>>>>>>>>
Load_list()	;������ʼ���б�
TrayTip,HotWindows,% "׼����ɿ�ʼʹ��`n��ǰ�汾�ţ�" UpdateInfo.Edition "`n����֧������rrsyycm@163.com",,1
Menu,Tray,Tip,% "HotWindows`n�汾:" UpdateInfo.Edition
;<<<<<<<<<<<<��Ҫѭ��>>>>>>>>>>>>
loop{
	WinGet,Wina_ID,ID,A
	Load_exe(Wina_ID)
	WinWaitNotActive,ahk_id %Wina_id%
	Load_exe(Wina_ID)
}
return

;<<<<<<<<<<<<��Ҫ���ܵı�ǩ>>>>>>>>>>>>
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
if (Boot="1") and (StrLens="1")	;���������뱣��ʲôҲû�з���
	Cancel()
if (Boot="2")	;û�п������뱣�������һ��
	Activate("1")
if (Boot="1") and (StrLens>"1")	;���������뱣������������
	Activate("1")
Cancel()
return

Add_exe:
	Gui,New
	Gui,Add_exe:New
	Gui,Add_exe:+LabelMyAdd +AlwaysOnTop
	Gui,Add_exe:Add,Text,xm,��ӳ����뽫�ļ����뱾����
	Gui,Add_exe:Add,ListView,xm w%ListWidth% vAdd_list r9,����|·��
	Gui,Add_exe:Add,Text,xm,�˴���ӳ���Ŀ¼c:\Users\*.exe��c:\Users\*.lnk
	Gui,Add_exe:Add,Edit,xm w%ListWidth% r5 vPath_list,%Path_list%
	Gui,Add_exe:Add,Button,xm Section gDele_exe,ɾ��ѡ�����(&D)
	Gui,Add_exe:Add,Button,ys gSubmit_exe,�������(&S)
	Gui,Add_exe:Show,,��ӳ����������б�
	Add_list()
return

Submit_exe:
	TrayTip,HotWindows,�ȴ��������,,1
	Gui,Submit,NoHide
	RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,Path_list,%Path_list%
	Load_list()
	Add_list()
	TrayTip,HotWindows,�������,,1
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
	TrayTip,HotWindows,ɾ�����,,3
	Gui,ListView,Hot_ListView
return

MyAddDropFiles:
	loop,Parse,A_GuiEvent,`n
		Add_quick(A_LoopField)
	Add_list()
	TrayTip,HotWindows,������,,1
return

Hot_ListView:
	EventInfo:=LV_GetNext()
return
;<<<<<<<<<<<<���ں���>>>>>>>>>>>>
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
				Level=��
			else
				Level=��
			LV_Add("Icon" . IL_Add(ImageListID,v.Path,1),k Level,v.Title)
		}
		LV_ModifyCol()
		SB_SetText("������ʷ��" . K_ThisHotkey . "")
		LV_Modify(1,"Select")
		LV_Modify(1,"Focus")
		GuiControl,+Redraw,MyListView
		Gui,Show,AutoSize CEnter,HotWindows
	}else{
		Tip_list:=K_ThisHotkey
		For k,v in WHERE_list
		{
			if v.Pid
				Level=��
			else
				Level=��
			Tip_list:=Tip_list "`n" k Level SubStr(v.Title,"1","25")
		}
		TrayTip,,%Tip_list%
	}
	if not K_ThisHotkey
		Cancel()
}

;<<<<<<<<<<<<��������>>>>>>>>>>>>
Load_list(){
	Suspend,On
	windowList =
	DetectHiddenWindows, Off ; makes DllCall("IsWindowVisible") unnecessary
	WinGet, windowList, List ; gather a list of Running programs
	loop, %windowList%
	{
		Load_exe(windowList%A_Index%)
		Progress,% Gers:=90//windowList*A_Index ,% Gers "%",������ǰ������Ϣ...,HotWindows
	}
	;�����ݿ�ɾ�������ڵĳ���
	Recordset := ComObjCreate("ADODB.Recordset")
	Recordset.CursorLocation:="3"
	Recordset.Open("SELECT Path FROM Quick","Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . Path_data . "")
	while !Recordset.EOF
	{
		Progress,% Gers:=90+10//Recordset.RecordCount*A_Index ,% Gers "%",ɾ�������ڳ���...,HotWindows
		Quick_Path:=Recordset.Fields["Path"].Value
		IfNotExist,%Quick_Path%
			SQL_Run("Delete FROM Quick WHERE Path='" Quick_Path "'")
		Recordset.MoveNext()
	}
	;ɾ������ǰ�ĳ�������
	SQL_Run("Delete FROM Activate WHERE Add_Time < '" A_Now - 172800 "'")
	;�����������
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
;<<<<<<<<<<<<MENU�Ĺ���>>>>>>>>>>>>
Traytip:
	if (Show_mode = "TrayTip")
		return
	RegRead,Bubble,HKEY_CURRENT_USER,SOFTWARE\Policies\Microsoft\Windows\Explorer,EnableLegacyBalloonNotifications
	if not Bubble {
		MsgBox,4,��Ҫ����,���'Yes'��Դ������������������Ҫ�����뱣����ٴ����á�
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
	Run "https://github.com/liumenggit/HotWindows#����������"
return

Auto:
	IfExist,%HotRun%
		RegDelete,HKEY_CURRENT_USER,Software\Microsoft\Windows\CurrentVersion\Run,HotRun
	else
		RegWrite,REG_SZ,HKEY_CURRENT_USER,Software\Microsoft\Windows\CurrentVersion\Run,HotRun,%A_ScriptFullPath%
	Menu,Tray,ToggleCheck,��������
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
	RegRead,Boot,HKEY_CURRENT_USER,HotWindows,Hotboot	;���뱣��
	if Boot=1
		Boot=2
	else
		Boot=1
	RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,Hotboot,%Boot%
	Menu,Tray,ToggleCheck,���뱣��
return

Dele_mdb:
	TrayTip,HotWindows,�ȴ��������,,1
	SQL_Run("Delete FROM Activate")
	SQL_Run("Delete FROM Nowlist")
	SQL_Run("Delete FROM Quick")
	RegDelete,HKEY_CURRENT_USER,HotWindows,HotStyles
	RegDelete,HKEY_CURRENT_USER,HotWindows,Path_list
	Load_list()
	TrayTip,HotWindows,�Ѿ�������м�¼,,1
return

Dele_mdb_Gui:
	TrayTip,HotWindows,�ȴ��������,,1
	SQL_Run("Delete FROM Activate")
	Load_list()
	TrayTip,HotWindows,�Ѿ�������ڼ�¼,,1
return

Dele_mdb_Exe:
	TrayTip,HotWindows,�ȴ��������,,1
	SQL_Run("Delete FROM Quick")
	RegDelete,HKEY_CURRENT_USER,HotWindows,Path_list
	Load_list()
	TrayTip,HotWindows,�Ѿ���������¼,,1
return


Reload:
	Reload
ExitApp:
	ExitApp

	;<<<<<<<<<<<<SQL����>>>>>>>>>>>>
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
SQL_Run(SQL){	;�����ݿ���������
	Recordset := ComObjCreate("ADODB.Recordset")
	Try Recordset.Open(SQL,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . Path_data . "")
		catch e
			return
}
SQL_Get(SQL){	;�����ݿ������������󷵻�
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

;���ߣ��������
;���ܣ�����GItHub��Ŀ���£�����ű��м���ʹ���Լ���GitHub��Ŀ��ַ
;���ܣ�����GitHub��Commitkey��ȡ�Ƿ����
;ע�⣺�ܹ�ʹ��GitHub������Ӧ�öԴ��붼�ǳ���Ϥ��ô��������Ҫ�������޸�

Git_Update(GitUrl,GressSet:="Hide"){
	if not W_InternetCheckConnection(GitUrl)
		return
	SplitPath,GitUrl,Project_Name
	RegRead,Reg_Commitkey,HKEY_CURRENT_USER,%Project_Name%,Commitkey
	if GressSet=Show
		Progress,100,% Reg_Commitkey " >>> " Git_CcommitKey.Edition,���������Ե�...,% Project_Name
	Git_CcommitKey:=Git_CcommitKey(GitUrl)
	if not Git_CcommitKey.Edition{	;��ȡ����ʧ�ܷ���
		Progress,Off
		return
	}
	if not Reg_Commitkey or (Reg_Commitkey<>Git_CcommitKey.Edition){	;���ڸ��¿�ʼ����
		Progress,1 T Cx0 FM10,��ʼ������,% Reg_Commitkey " >>> " Git_CcommitKey.Edition " ��飺" Git_CcommitKey.Commit,% Project_Name
		Git_Downloand(Git_CcommitKey,Project_Name)
	}else{
		Progress,,,���޸���,% Project_Name
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

SmartZip(s, o, t = 16)	;���ý�ѹ����
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

W_InternetCheckConnection(lpszUrl){ ;���FTP�����Ƿ������
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
				fj	?(DllCall("Wininet.dll\InternetReadFile", "Ptr", F, "Ptr", &Buf + C, "UInt", S, "UIntP", R),C += R,DllCall("QueryPerformanceCounter", "Int64*", y),((t:=(y-x)/i) >=1)?(Test(e,Round(c/fx,2) fz " | " Round(((c-w)/1024)/t) "KB/��",Round(c/fx/percent*100)),x:=y,w:=c):"")
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
