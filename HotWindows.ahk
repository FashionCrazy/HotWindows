DetectHiddenWindows,On
#WinActivateForce
#Persistent
#SingleInstance force
#UseHook
#InstallKeybdHook
#ErrorStdOut
ComObjError(false)
SetBatchLines -1

;<<<<<<<<<<<<Ĭ��ֵ>>>>>>>>>>>>
Path_data=%A_ScriptDir%\HotWindows.mdb	;���ݿ��ַ

;<<<<<<<<<<<<WIN10 WIN8����Ҫ������ֵ>>>>>>>>>>>>
RegRead,Bubble,HKEY_CURRENT_USER,SOFTWARE\Policies\Microsoft\Windows\Explorer,EnableLegacyBalloonNotifications
if not Bubble
	MsgBox,4,��Ҫ����,�ű���Ҫʹ��������ʾ���Yesȷ���л�Ϊ������ʾ`n����ָ��������������������и���
		IfMsgBox Yes
		{
			RegWrite,REG_DWORD,HKEY_CURRENT_USER,SOFTWARE\Policies\Microsoft\Windows\Explorer,EnableLegacyBalloonNotifications,1
			RunWait %comspec% /c "taskkill /f /im explorer.exe",,Hide
			Run %comspec% /c "start c:\Windows\explorer.exe",,Hide
		}

Progress,,��ʼ��,��ʼ�����Ե�...,HotWindows
;<<<<<<<<<<<<Ԥ��������>>>>>>>>>>>>
Show_modes=TrayTip,ListView
Hot_keys=Space,Tab
Loop,parse,Show_modes,`,
	Menu,Show_mode,Add,%A_LoopField%,Show_mode
Loop,parse,Hot_keys,`,
	Menu,Hot_key,Add,%A_LoopField%,Hot_key
Menu,Tray,Add,��������,Auto
Menu,Tray,Add,������ʾ,Bubble
Menu,Tray,Add,���뱣��,Boot
Menu,Tray,Add
Menu,Tray,Add,��ʾ��ʽ,:Show_mode
Menu,Tray,Add,�����ȼ�,:Hot_key
Menu,Dele_mdb,Add,������ڼ�¼,Dele_mdb_Gui
Menu,Dele_mdb,Add,��������¼,Dele_mdb_Exe
Menu,Dele_mdb,Add,�����ʽ��¼,Dele_mdb_Style
Menu,Dele_mdb,Add,������м�¼,Dele_mdb
Menu,Tray,Add,�����¼,:Dele_mdb
Menu,Tray,Add,��ӳ���,Add_exe
Menu,Tray,Add
Menu,Tray,Add,�����ű�,Reload
Menu,Tray,Add,�˳��ű�,ExitApp
Menu,Tray,Icon,��ʾ��ʽ,MenuIco.icl,6
Menu,Tray,Icon,�����ȼ�,MenuIco.icl,4
Menu,Tray,Icon,��ӳ���,MenuIco.icl,1
Menu,Tray,Icon,�����¼,MenuIco.icl,2
Menu,Tray,Icon,�����ű�,MenuIco.icl,5
Menu,Tray,Icon,�˳��ű�,MenuIco.icl,3
Menu,Tray,NoStandard
Menu,Tray,Icon,MenuIco.icl,7
Menu,Tray,Icon,,,1
RegRead,HotRun,HKEY_CURRENT_USER,Software\Microsoft\Windows\CurrentVersion\Run,HotRun
RegRead,Bubble,HKEY_CURRENT_USER,SOFTWARE\Policies\Microsoft\Windows\Explorer,EnableLegacyBalloonNotifications
RegRead,Show_mode,HKEY_CURRENT_USER,HotWindows,HotShow_mode	;��ʾ��ʽ
RegRead,Hot_Set_key,HKEY_CURRENT_USER,HotWindows,HotHot_key	;�����ȼ�
RegRead,Boot,HKEY_CURRENT_USER,HotWindows,Hotboot	;���뱣��
RegRead,Styles,HKEY_CURRENT_USER,HotWindows,HotStyles	;��ʽ�б�
RegRead,Path_list,HKEY_CURRENT_USER,HotWindows,Path_list	;�Զ��������ӹ���
if not Path_list
	RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,Path_list,%A_Desktop%\*.lnk`n
RegRead,Path_list,HKEY_CURRENT_USER,HotWindows,Path_list	;�Զ��������ӹ���
if not Styles{
	Styles=0x14EF0000,0x15CF0000,0x34CF0000,0x860F0000,0x860E0000,0x36CF0000,0x17CF0000,0x84C80000,0xB4CF0000,0x94CA0000,0x95CF0000,0x94CF0000,0x94000000
	RegWrite,REG_MULTI_SZ,HKEY_CURRENT_USER,HotWindows,HotStyles,%Styles%
}
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
	Menu,Show_mode,ToggleCheck,ListView
}else{
	Menu,Show_mode,ToggleCheck,%Show_mode%
}
IfExist,%HotRun%
{
	RegWrite,REG_SZ,HKEY_CURRENT_USER,Software\Microsoft\Windows\CurrentVersion\Run,HotRun,%A_ScriptFullPath%
	Menu,Tray,ToggleCheck,��������
}
if Bubble
	Menu,Tray,ToggleCheck,������ʾ
if not Boot{
	Boot:=1
	RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,Hotboot,% Boot
	Menu,Tray,ToggleCheck,���뱣��
}else if (Boot="1"){
	Menu,Tray,ToggleCheck,���뱣��
}

;<<<<<<<<<<<<�ȼ�����>>>>>>>>>>>>
Layout=qwertyuiopasdfghjklzxcvbnm
Loop,Parse,Layout
{
	Layouts:=A_LoopField
	Loop,parse,Hot_keys,`,
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
global Styles,Path_data,Show_mode,K_ThisHotkey,WHERE_list,Path_list,Ger,Gers,Starts,NewEdition

;<<<<<<<<<<<<������>>>>>>>>>>>>
UpdateInfo:=Git_Update("https://github.com/liumenggit/HotWindows","Show")

;<<<<<<<<<<<<DLL����>>>>>>>>>>>>
tcmatch := "Tcmatch.dll"
hModule := DllCall("LoadLibrary", "Str", tcmatch, "Ptr")

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
	SQL_Run("CREATE TABLE Now_list(Title varchar(255),Pid varchar(255),Path varchar(255),GetStyle varchar(255))")	;��ӳ������ݿ��
	SQL_Run("CREATE TABLE Activate(Title varchar(255),Times varchar(255))")	;��ӳ������ݿ��
	SQL_Run("CREATE TABLE Quick(Title varchar(255),Pid varchar(255),Path varchar(255),GetStyle varchar(255))")	;��ӳ������ݿ��
}else{
	SQL_Run("DELETE FROM Now_list")
}

;<<<<<<<<<<<<�����б�>>>>>>>>>>>>
Load_list()	;������ʼ���б�
TrayTip,HotWindows,% "׼����ɿ�ʼʹ��`n��ǰ�汾�ţ�" UpdateInfo.Edition "`n����֧������rrsyycm@163.com`n" NewCnzz,,1
Menu,Tray,Tip,% "HotWindows`n�汾:" UpdateInfo.Edition
;<<<<<<<<<<<<��Ҫѭ��>>>>>>>>>>>>
loop{
	WinGet,Wina_ID,ID,A
	WinGet,Exe_Name,ProcessName,ahk_id %Wina_id%
	WinGet,Get_Style,Style,ahk_id %Wina_id%
	if Get_Style not in %Styles%
	{
		Styles=%Styles%`,%Get_Style%
		RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,HotStyles,% Styles
	}
	Load_exe(Exe_Name)
	WinWaitNotActive,ahk_id %Wina_id%
	Load_exe(Exe_Name)
}
Return

;<<<<<<<<<<<<��Ҫ���ܵı�ǩ>>>>>>>>>>>>
Layout:
	Critical
	StringRight,H_ThisHotkey,A_ThisHotkey,1
	K_ThisHotkey:=K_ThisHotkey H_ThisHotkey
	StrLens := StrLen(K_ThisHotkey)
	ToolTip,,%K_ThisHotkey%
	if (StrLens="1"){
		loop,9{
			Hotkey,%A_Index%,Table
			Hotkey,%A_Index%,On
		}
		SetTimer,Key_wait,1
	}
	SQL_List("SELECT Activate.title,Activate.times,t1.pid,t1.path,t1.getstyle FROM Activate LEFT JOIN (SELECT * FROM Now_list UNION SELECT * FROM Quick) AS t1 ON Activate.title = t1.title WHERE t1.pid IS NOT NULL OR t1.path IS NOT NULL ORDER BY Activate.Times +- 1 DESC,t1.GetStyle DESC",K_ThisHotkey)
	if WHERE_list.Length() and K_ThisHotkey{
		Show_list(WHERE_list)
		if (WHERE_list.Length()="1"){
			Activate("1")
			Send {%Hot_Set_key% Up}
		}
	}else{
		Critical off
		Cancel()
	}
Return


Table:
	Activate(A_ThisHotkey)
Return

Key_wait:
	SetTimer,Key_wait,off
	KeyWait,%Hot_Set_key%,L
	if not K_ThisHotkey{
		Cancel()
		Return
	}
	if (EventInfo<>"0") and (Show_mode="ListView"){
		if (Boot="1") and (StrLens="1")
			Cancel()
		if (Boot="2")
			Activate(EventInfo)
		if (Boot="1") and (StrLens>"1")
			Activate(EventInfo)
		Return
	}
	if (Boot="1") and (StrLens="1")	;���������뱣��ʲôҲû�з���
		Cancel()
	if (Boot="2")	;û�п������뱣�������һ��
		Activate("1")
	if (Boot="1") and (StrLens>"1")	;���������뱣������������
		Activate("1")
	Cancel()
Return

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
Return

Submit_exe:
	TrayTip,HotWindows,�ȴ��������,,1
	Gui,Submit,NoHide
	RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,Path_list,%Path_list%
	Load_list()
	Add_list()
	TrayTip,HotWindows,�������,,1
Return

Dele_exe:
Gui,ListView,Add_list
RowNumber=0
Loop
{
    RowNumber:=LV_GetNext(RowNumber)
    if not RowNumber
        break
    LV_GetText(dPath,RowNumber,2)
	SQL_Run("DELETE FROM Quick WHERE Path='" dPath "'")
}
Add_list()
TrayTip,HotWindows,ɾ�����,,3
Gui,ListView,Hot_ListView
Return

MyAddDropFiles:
Loop,Parse,A_GuiEvent,`n
	Add_quick(A_LoopField)
Add_list()
TrayTip,HotWindows,������,,1
Return

Hot_ListView:
EventInfo:=LV_GetNext()
Return
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
			SQL_Run("DELETE FROM Quick WHERE Path='" Recordset.Fields["Path"].Value "'")
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
	if Activate{
		WinActivate,ahk_id %Activate%
	}else{
		Try RunWait %Path%
		catch e
			Return
	}
	if not Sql_Get("SELECT Times FROM Activate WHERE Title='" Title "'")
		SQL_Run("Insert INTO Activate (Title,Times) VALUES ('" Title "','1')")
	else
		SQL_Run("UPDATE Activate SET Times = Times+1 WHERE Title='" Title "'")
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
		Gui,Show,AutoSize Center,HotWindows
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
	DetectHiddenWindows,Off
	WinGet,ID_list,List,,,Program Manager
	DetectHiddenWindows,On
	Ger:=100//ID_list
	Gers:=0
	loop,%ID_list%
	{
		This_id := ID_list%A_Index%
		WinGet,Exe_Name,ProcessName,ahk_id %This_id%
		IfNotInString,Exe_Names,%Exe_Name%
			Load_exe(Exe_Name)
		else
			Gers+=100//ID_list
		Exe_Names=%Exe_Name%`n%Exe_Names%
	}
	Starts=Yes
	Recordset := ComObjCreate("ADODB.Recordset")
	Recordset.Open("SELECT Path FROM Quick","Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . Path_data . "")
	while !Recordset.EOF
	{
		Quick_Path:=Recordset.Fields["Path"].Value
		;if not GetIconCount(Quick_Path)
		IfNotExist,%Quick_Path%
			SQL_Run("DELETE FROM Quick WHERE Path='" Quick_Path "'")
		Recordset.MoveNext()
	}
	loop,Parse,Path_list,`n
		loop,%A_LoopField%
			Add_quick(A_LoopFileLongPath)
	Add_list()
    Progress,100
	Progress,Off
	Suspend,Off
}
Load_exe(Exe_Name){
	if not Exe_Name
		Return
	SQL_Run("DELETE FROM Now_list WHERE Path LIKE '%" Exe_Name "'")
	WinGet,WinList,List,ahk_exe %Exe_Name%
	WinGet,Path,ProcessPath,ahk_exe %Exe_Name%
	SplitPath,Path,OutFileName,OutDir,OutExtension,OutNameNoExt,OutDrive
	;Windowsvar=C:\Windows\
	Add_quick(Path)
	loop,%WinList% {
		PID:=WinList%A_Index%
		WinGet,GetStyle,Style,ahk_id %PID%
		WinGetTitle,Title,ahk_id %PID%
		WinGet,Path,ProcessPath,ahk_id %PID%
		if not Starts and Title
    		Progress,% Gers+=Ger/WinList ,% Title,������ǰ������Ϣ...,HotWindows
		if GetStyle in %Styles%
			if Title and GetIconCount(Path){
				if Sql_Get("SELECT COUNT(*) FROM Now_list WHERE Title='" Title "'")
					continue
				SQL_Run("Insert INTO Now_list (Title,PID,Path,GetStyle) VALUES ('" Title "','" PID "','" Path "','" GetStyle "')")
				if not Sql_Get("SELECT Times FROM Activate WHERE Title='" Title "'")
					SQL_Run("Insert INTO Activate (Title,Times) VALUES ('" Title "','1')")
			}
	}
}

Add_quick(Path){
	IfNotInString,Path,%A_WinDir%
	{
		SplitPath,Path,OutFileName,OutDir,OutExtension,OutNameNoExt,OutDrive
			if not Sql_Get("SELECT COUNT(*) FROM Quick WHERE Path='" Path "'"){
				SQL_Run("DELETE FROM Quick WHERE Path='" Path "'")
				SQL_Run("Insert INTO Quick (Title,Path) VALUES ('" OutNameNoExt "','" Path "')")
				if not Sql_Get("SELECT Times FROM Activate WHERE Title='" OutNameNoExt "'")
					SQL_Run("Insert INTO Activate (Title,Times) VALUES ('" OutNameNoExt "','1')")
			}
	}
}

GetIconCount(file){
	Menu, test, add, test, handle
	Loop
	{
		try {
			id++
			Menu, test, Icon, test, % file, % id
		} catch error {
			break
		}
	}
Return id-1
}
handle:
Return
;<<<<<<<<<<<<MENU�Ĺ���>>>>>>>>>>>>
Bubble:
if Bubble
	RegWrite,REG_DWORD,HKEY_CURRENT_USER,SOFTWARE\Policies\Microsoft\Windows\Explorer,EnableLegacyBalloonNotifications,0
else
	RegWrite,REG_DWORD,HKEY_CURRENT_USER,SOFTWARE\Policies\Microsoft\Windows\Explorer,EnableLegacyBalloonNotifications,1
	Menu,Tray,ToggleCheck,������ʾ
	RunWait %comspec% /c "taskkill /f /im explorer.exe",,Hide
	Run %comspec% /c "start c:\Windows\explorer.exe",,Hide
Return

Auto:
	IfExist,%HotRun%
		RegDelete,HKEY_CURRENT_USER,Software\Microsoft\Windows\CurrentVersion\Run,HotRun
	else
		RegWrite,REG_SZ,HKEY_CURRENT_USER,Software\Microsoft\Windows\CurrentVersion\Run,HotRun,%A_ScriptFullPath%
	Menu,Tray,ToggleCheck,��������
Return

Show_mode:
	Show_mode:=A_ThisMenuItem
	Loop,parse,Show_modes,`,
		Menu,Show_mode,Uncheck,%A_LoopField%
	Menu,Show_mode,ToggleCheck,%A_ThisMenuItem%
	RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,HotShow_mode,%A_ThisMenuItem%
Return
Hot_key:
	Loop,Parse,Layout
	{
		Hotkey,~%Hot_Set_key% & %A_LoopField%,off
		Hotkey,~%A_ThisMenuItem% & %A_LoopField%,On
	}
	Hot_Set_key:=A_ThisMenuItem
	Loop,parse,Hot_keys,`,
		Menu,Hot_key,Uncheck,%A_LoopField%
	Menu,Hot_key,ToggleCheck,%Hot_Set_key%
	RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,HotHot_key,%A_ThisMenuItem%
Return

Boot:
RegRead,Boot,HKEY_CURRENT_USER,HotWindows,Hotboot	;���뱣��
if Boot=1
	Boot=2
else
	Boot=1
RegWrite,REG_SZ,HKEY_CURRENT_USER,HotWindows,Hotboot,%Boot%
Menu,Tray,ToggleCheck,���뱣��
Return

Dele_mdb:
	TrayTip,HotWindows,�ȴ��������,,1
	SQL_Run("DELETE FROM Activate")
	SQL_Run("DELETE FROM Now_list")
	SQL_Run("DELETE FROM Quick")
	RegDelete,HKEY_CURRENT_USER,HotWindows,HotStyles
	RegDelete,HKEY_CURRENT_USER,HotWindows,Path_list
	Load_list()
	TrayTip,HotWindows,�Ѿ�������м�¼,,1
Return

Dele_mdb_Gui:
	TrayTip,HotWindows,�ȴ��������,,1
	SQL_Run("DELETE FROM Activate")
	Load_list()
	TrayTip,HotWindows,�Ѿ�������ڼ�¼,,1
Return

Dele_mdb_Exe:
	TrayTip,HotWindows,�ȴ��������,,1
	SQL_Run("DELETE FROM Quick")
	RegDelete,HKEY_CURRENT_USER,HotWindows,Path_list
	Load_list()
	TrayTip,HotWindows,�Ѿ���������¼,,1
Return

Dele_mdb_Style:
	TrayTip,HotWindows,�ȴ��������,,1
	RegDelete,HKEY_CURRENT_USER,HotWindows,HotStyles
	Load_list()
	TrayTip,HotWindows,�Ѿ������ʽ��¼,,1
Return

Reload:
	Reload
ExitApp:
	ExitApp

;<<<<<<<<<<<<SQL����>>>>>>>>>>>>
SQL_List(SQL,K_ThisHotkey){
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
				SQL_Run("DELETE FROM Now_list WHERE PID='" Recordset.Fields["PID"].Value "'")
				Recordset.MoveNext()
				Continue
			}
		}
		else if wPath
		{
			IfNotExist,%wPath%
			{
				SQL_Run("DELETE FROM Now_list WHERE Path='" Recordset.Fields["Path"].Value "'")
				Recordset.MoveNext()
				Continue
			}
		}
		if (matched := DllCall("Tcmatch\MatchFileW","WStr",K_ThisHotkey,"WStr",Recordset.Fields["Title"].Value)){
				WHERE_time++
				WHERE_list[WHERE_time]:={Title:Recordset.Fields["Title"].Value,PID:Recordset.Fields["PID"].Value,Path:Recordset.Fields["Path"].Value,GetStyle:Recordset.Fields["GetStyle"].Value,Times:Recordset.Fields["Times"].Value}
		}
		Recordset.MoveNext()
	}
	Return
}
SQL_Run(SQL){	;�����ݿ���������
	Recordset := ComObjCreate("ADODB.Recordset")
	Try Recordset.Open(SQL,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . Path_data . "")
	catch e
	Return
}
SQL_Get(SQL){	;�����ݿ������������󷵻�
	Recordset := ComObjCreate("ADODB.Recordset")
	Try Recordset.Open(SQL,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . Path_data . "")
	catch e
		Return 0
	Try Return Recordset.Fields[0].Value
	catch e
		Return 0
}

;���ߣ��������
;���ܣ�����GItHub��Ŀ���£�����ű��м���ʹ���Լ���GitHub��Ŀ��ַ
;���ܣ�����GitHub��Commitkey��ȡ�Ƿ����
;ע�⣺�ܹ�ʹ��GitHub������Ӧ�öԴ��붼�ǳ���Ϥ��ô��������Ҫ�������޸�

Git_Update(GitUrl,GressSet:="Hide"){
	if not W_InternetCheckConnection(GitUrl)
		Return
	SplitPath,GitUrl,Project_Name
	RegRead,Reg_Commitkey,HKEY_CURRENT_USER,%Project_Name%,Commitkey
	if GressSet=Show
		Progress,100,% Reg_Commitkey " >>> " Git_CcommitKey.Edition,���������Ե�...,% Project_Name
	Git_CcommitKey:=Git_CcommitKey(GitUrl)
	if not Git_CcommitKey.Edition{	;��ȡ����ʧ�ܷ���
		Progress,Off
		Return
	}
	if not Reg_Commitkey or (Reg_Commitkey<>Git_CcommitKey.Edition){	;���ڸ��¿�ʼ����
		Progress,1 T Cx0 FM10,��ʼ������,% Reg_Commitkey " >>> " Git_CcommitKey.Edition " ��飺" Git_CcommitKey.Commit,% Project_Name
		Git_Downloand(Git_CcommitKey,Project_Name)
	}else{
		Progress,,,���޸���,% Project_Name
	}
	Progress,Off
	Return Git_CcommitKey
}

Git_Downloand(DownloandInfo,Project_Name){
	DownUrl:="https://github.com" DownloandInfo.Down
	SplitPath,A_ScriptName,,,,A_name
	SplitPath,DownUrl,DownName,,,OutNameNoExt
	if not Z_Down(DownUrl,"",A_name,A_Temp "\" DownName){
		Progress,Off
		Return
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
		Return, -1
	oShell := ComObjCreate("Shell.Application")
	if InStr(FileExist(o), "D") or (!FileExist(o) and (SubStr(s, -3) = ".zip"))
	{
		if !o
			o := A_ScriptDir
		else ifNotExist, %o%
			FileCreateDir, %o%
		Loop, %o%, 1
			sObjectLongName := A_LoopFileLongPath
		oObject := oShell.NameSpace(sObjectLongName)
		Loop, %s%, 1
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
		Return {Edition:NewEdition,Down:Downloand,Commit:Committitle}
	}catch e {
		Return
	}
}

W_InternetCheckConnection(lpszUrl){ ;���FTP�����Ƿ������
	FLAG_ICC_FORCE_CONNECTION := 0x1
	dwReserved := 0x0
	Return, DllCall("Wininet.dll\InternetCheckConnection", "Ptr", &lpszUrl, "UInt", FLAG_ICC_FORCE_CONNECTION, "UInt", dwReserved, "Int")
}
Z_Down(url:="http://61.135.169.125/forbiddenip/forbidden.html", Proxy:="",e:="utf-8", File:="",byref buf:=""){
	if (!(File?o:=FileOpen(File, "w"):1) or !DllCall("LoadLibrary", "str", "wininet") or !(h := DllCall("wininet\InternetOpen", "str", "", "uint", Proxy?3:1, "str", Proxy, "str", "", "uint", 0)))
		Return 0
	c:=s:=0
	if (f := DllCall("wininet\InternetOpenUrl", "ptr", h, "str", url, "ptr", 0, "uint", 0, "uint", 0x80003000, "ptr", 0, "ptr"))
		{
			if File or IsByRef(buf)
			{
				VarSetCapacity(buffer,1024,0),VarSetCapacity(bufferlen,4,0)
				Loop, 5
				if (DllCall("wininet\HttpQueryInfo","uint",f, "uint", 22, "uint", &buffer, "uint", &bufferlen, "uint", 0) = 1)
				{
					Progress,+20
					y:= Trim(StrGet(&buffer)," `r`n"),q:=[]
					Loop,parse,y,`r`n
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
