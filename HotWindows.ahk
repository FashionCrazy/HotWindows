DetectHiddenWindows,On
#Persistent
#SingleInstance force
#UseHook
#InstallKeybdHook
#Include %A_ScriptDir%\JSON.ahk
RegRead,Bubble,HKEY_CURRENT_USER,SOFTWARE\Policies\Microsoft\Windows\Explorer,EnableLegacyBalloonNotifications
if not Bubble
	MsgBox,4,��Ҫ����,�ű���Ҫʹ��������ʾ���Yesȷ���л�Ϊ������ʾ`n����ָ��������������и���
		IfMsgBox Yes
			RegWrite,REG_DWORD,HKEY_CURRENT_USER,SOFTWARE\Policies\Microsoft\Windows\Explorer,EnableLegacyBalloonNotifications,1
GuiArr := Object()
Edition:=201706
RegRead,LastTime,HKEY_LOCAL_MACHINE,SOFTWARE\TestKey,HotEdit
RegWrite,REG_SZ,HKEY_LOCAL_MACHINE,SOFTWARE\TestKey,HotEdit,%Edition%
if LastTime and (LastTime<Edition){
	TrayTip,�����ɹ�,�Ѵ�%LastTime%������%Edition%,,1
	Sleep,2000
}
if W_InternetCheckConnection("https://github.com"){
	TrayTip,������,ȷ��������ͨ,,1
	GetJson:=JSON.load(Update("http://autoahk.com/hotwindows.php"))
	if (GetJson[1].time>Edition){
		Time:=GetJson[1].time
		Inf:=GetJson[1].inf
		Url:=GetJson[1].URL
		RunWait https://github.com/liumenggit/HotWindows#������ʷ
		MsgBox,4,�汾����,�Ƿ���µ����°汾��%Time%`n%Inf%
		IfMsgBox Yes
			gosub,Downloand
	}
}else{
	TrayTip,������,����δ����,,3
}
Menu,Tray,Add,Hot-Windows,Menu_show
Menu,Tray,Add,��������,Auto
Menu,Tray,Add,������ʾ,Bubble
Menu,Tray,Add,�����ű�,Reload
Menu,Tray,Add,�˳��ű�,ExitApp
Menu,Tray,Default,Hot-Windows
Menu,Tray,NoStandard

Gui,PS:+HwndMyGuiHwnd -MaximizeBox -MinimizeBox
Gui,PS:Add,Text,,�ȼ���
Gui,PS:Add,DDL,vDDL1 AltSubmit
Gui,PS:Add,Text,,������
Gui,PS:Add,DDL,vDDL2 AltSubmit
Gui,PS:Add,Text,,��ʾ��ʽ
Gui,PS:Add,DDL,vDDL3 AltSubmit
Gui,PS:Add,Text,,��С���ȼ�
Gui,PS:Add,Hotkey,vWinmin
Gui,PS:Add,Text,,����ȼ�
Gui,PS:Add,Hotkey,vWinmax
Gui,PS:Add,Text,,�����ȼ�
Gui,PS:Add,Hotkey,vWinmove
Gui,PS:Add,Button,gSubmit w135,��������
Gui,+AlwaysOnTop +ToolWindow -Caption -MaximizeBox -MinimizeBox
Gui,Add,ListView,w600 R9,ID|Title
RegRead,Bubble,HKEY_CURRENT_USER,SOFTWARE\Policies\Microsoft\Windows\Explorer,EnableLegacyBalloonNotifications
RegRead,HotRun,HKEY_CURRENT_USER,Software\Microsoft\Windows\CurrentVersion\Run,HotRun
if Bubble
	Menu,Tray,ToggleCheck,������ʾ
IfExist,%HotRun%
{
	RegWrite,REG_SZ,HKEY_CURRENT_USER,Software\Microsoft\Windows\CurrentVersion\Run,HotRun,%A_ScriptFullPath%
	Menu,Tray,ToggleCheck,��������
}
tcmatch := "Tcmatch.dll"
hModule := DllCall("LoadLibrary", "Str", tcmatch, "Ptr")
Gui_Submit("0")
Sleep,199
Hot:=GuiArr["DDL1"]
Key:=GuiArr["DDL2"]
DDL:=GuiArr["DDL3"]
Arrays:=GetArray()
Layout=qwertyuiopasdfghjklzxcvbnm
Loop,Parse,Layout
	Hotkey,%A_LoopField%,Layout
TrayTip,HotWindows,�����Ѿ�����׼��`n�������ͼ������`n��ǰ�汾��%Edition%,,1
loop{
	WinGet,WinGet_ID,ID,A
	WinGet,getexe,ProcessName,ahk_id %WinGet_ID%
	WinGet,WinList,List,ahk_exe %getexe%
	loop,%WinList% {
		id:=WinList%A_Index%
		WinGetTitle,Title,ahk_id %id%
		if Title
			Arrays[Title] := id
	}
	WinWaitNotActive,ahk_id %WinGet_ID%
	For k,v in Arrays{
		IfWinNotExist,ahk_id %v%
			Arrays.Delete(k)
		IfWinNotExist,%k%
			Arrays.Delete(k)
	}
}
return

Layout:
if GetKeyState(key,"P"){
	;if (A_TimeSincePriorHotkey<"200") and (ThisHotkey=A_ThisHotkey){
	StringReplace,ThisHotkey,A_ThisHotkey,~
	RegRead,boss_exe,HKEY_LOCAL_MACHINE,SOFTWARE\TestKey,bossexe%ThisHotkey%
	RegRead,boss_path,HKEY_LOCAL_MACHINE,SOFTWARE\TestKey,bosspath%ThisHotkey%
	DetectHiddenWindows,Off
	if boss_exe
	{
		IfWinActive,ahk_exe %boss_exe%
		{
			;WinMinimize,ahk_exe %boss_exe%
			RegDelete,HKEY_LOCAL_MACHINE,SOFTWARE\TestKey,bossexe%ThisHotkey%
			RegDelete,HKEY_LOCAL_MACHINE,SOFTWARE\TestKey,bosspath%ThisHotkey%
			TrayTip,,����%boss_exe%`n�ȼ���%A_ThisHotkey%`n��ɾ��,,1				
		}else{
			IfWinExist,ahk_exe %boss_exe%
				WinActivate,ahk_exe %boss_exe%
			else
				Run,%boss_path%,,,RunPid
		}
	}else{
		WinGet,boss_exe,ProcessName,A
		WinGet,boss_path,ProcessPath,A
		RegWrite,REG_SZ,HKEY_LOCAL_MACHINE,SOFTWARE\TestKey,bossexe%ThisHotkey%,%boss_exe%
		RegWrite,REG_SZ,HKEY_LOCAL_MACHINE,SOFTWARE\TestKey,bosspath%ThisHotkey%,%boss_path%
		TrayTip,,����%boss_exe%`n�ȼ���%A_ThisHotkey%,,1
		WinMinimize A
	}
	DetectHiddenWindows,On
	;}
}
if GetKeyState(Hot,"P"){
	StringReplace,ThisHotkey,A_ThisHotkey,~
	Hots = %Hots%%ThisHotkey%
	vars := StrLen(Hots)
	lists := Object()
	listv := Object()
	marry :=
	if vars=1
		loop,9{
			Hotkey,%A_Index%,Table
			Hotkey,%A_Index%,On
		}
	for k,v in Arrays{
		if (matched := DllCall("Tcmatch\MatchFileW","WStr",Hots,"WStr",k)){
			marry++
			lists[k]:=v
			listv[marry]:=v
		}
	}
	if marry{
		ListViewAdd(lists)
		if (marry="1")
			Activate("1")
	}
	if vars=1
		goto,WaitHot
}else{
	if GetKeyState("CapsLock","T")
		StringUpper,ThisHotkey,A_ThisHotkey
	else
		ThisHotkey:=A_ThisHotkey
	Send %ThisHotkey%
	Cancel()
	Hots:=
}
	ThisHotkey:=A_ThisHotkey
return

ListViewAdd(ls){
	global DDL
	if DDL=ListView
	{
		GuiControl,-Redraw,MyListView
		LV_Delete()
		WinAll:= ls.MaxIndex()
		ImageListID := IL_Create(WinAll)
		LV_SetImageList(ImageListID)
		For k,v in ls
		{
			WinGet,Route,ProcessPath,ahk_id %v%
			LV_Add("Icon" . IL_Add(ImageListID,Route,1),A_Index,k)
		}
		LV_ModifyCol()
		GuiControl,+Redraw,MyListView
		Gui,Show
	}else{
		For k,v in ls
			list=%list%`n%A_index%-%k%
		list := Trim(list,"`n")
		TrayTip,,%list% ; `n %A_ThisHotkey% `n %hots% `n %vars% `n %marry%
	}
}

WaitHot:
	KeyWait,%Hot%,L
	if Hots
		Activate("1")
return

Table:
	Activate(A_ThisHotkey)
return

WinMinimize:
	WinMinimize A
return

WinMaximize:
	WinMaximize A
return

WinMove:
	D_Width:=(A_ScreenWidth/2)
	D_Height:=(A_ScreenHeight/2)
	WinGetPos,X,Y,Width,Height,A
	WinMove,A,,A_ScreenWidth/8,A_ScreenHeight/8,(A_ScreenWidth/8)*6,(A_ScreenHeight/8)*6
return

Bubble:
if Bubble
	RegWrite,REG_DWORD,HKEY_CURRENT_USER,SOFTWARE\Policies\Microsoft\Windows\Explorer,EnableLegacyBalloonNotifications,0
else
	RegWrite,REG_DWORD,HKEY_CURRENT_USER,SOFTWARE\Policies\Microsoft\Windows\Explorer,EnableLegacyBalloonNotifications,1
	Menu,Tray,ToggleCheck,������ʾ
return

Auto:
	IfExist,%HotRun%
		RegDelete,HKEY_CURRENT_USER,Software\Microsoft\Windows\CurrentVersion\Run,HotRun
	else
		RegWrite,REG_SZ,HKEY_CURRENT_USER,Software\Microsoft\Windows\CurrentVersion\Run,HotRun,%A_ScriptFullPath%
	Menu,Tray,ToggleCheck,��������
return

Reload:
	Reload
ExitApp:
	ExitApp

Submit:
Gui,Submit,NoHide
if (DDL1=DDL2){
	TrayTip,HotWindows,���������ȼ����ȼ�������ͬ,,3
	return
}
Gui_Submit("1")
Gui_Submit("0")
Hot:=GuiArr["DDL1"]
Key:=GuiArr["DDL2"]
DDL:=GuiArr["DDL3"]
TrayTip,HotWindows,�����Ѽ�¼,,1
return

Menu_show:
DetectHiddenWindows,Off
IfWinNotExist,ahk_id %MyGuiHwnd%
	Gui,PS:Show
else
	Gui,PS:Cancel
DetectHiddenWindows,On
return

Downloand:
	Gui,Add,Text,xm ym w233 vLabel1,���ڳ�ʼ��...
	Gui,Add,Text,xm y24 w140 vLabel2,
	Gui,Add,Text, x150 y24 w80 vLabel3,
	Gui,Add,Button, x260 y10 w50 h25 gCancel, ȡ��
	Gui,Add,Progress, x10 y45 w300 h20 vMyProgress -Smooth
	Gui, +ToolWindow +AlwaysOnTop
	SysGet, m, MonitorWorkArea,1
	x:=A_ScreenWidth-520
	y:=A_ScreenHeight-180
	Gui,Show,w320 x%x% y%y% , �ļ�����
	Gui +LastFound
	SplitPath, URL, FN,,,, DN
	FN:=(FN ? FN : DN)
	SAVE=%A_ScriptDir%\HotWindows-master.zip
	DllCall("QueryPerformanceCounter", "Int64*", T1)
	WP1=0
	T2=0
	WP2=0
	if ((E:=InternetFileRead( binData, URL, False, 1024)) > 0 && !ErrorLevel)
	{
		VarZ_Save(binData, SAVE)
		GuiControl, Text, Label1, ������ɡ�
		Sleep, 500
		FileCopyDir,%A_ScriptDir%,%A_ScriptDir%\history\%Edition%,1
		D_history=%A_ScriptDir%\history\%time%
		FileCreateDir,%D_history%
		SmartZip(SAVE,D_history)
		FileDelete,%SAVE%
		gosub,ExitSub
		ExitApp
	}else{
		ERR := (E<0) ? "����ʧ�ܣ��������Ϊ" . E : "���ع����г���δ��������ء����ֶ����¡�"
		GuiControl, Text, Label1, %ERR%
		Sleep, 500
		Gui,Destroy
		return
	}
	DllCall( "FreeLibrary", UInt,DllCall( "GetModuleHandle", Str,"wininet.dll") )
return

ExitSub:
bat=
		(LTrim
:start
	ping 127.0.0.1 -n 2>nul
	del `%1
	if exist `%1 goto start
	xcopy %D_history%\HotWindows-master %A_ScriptDir% /s/e/y
	rd /s/q %D_history%
	start %A_ScriptFullPath%
	del `%0
	)
	batfilename=Delete.bat
	IfExist %batfilename%
		FileDelete %batfilename%
	FileAppend, %bat%, %batfilename%
	Run,%batfilename% , , Hide
	ExitApp
return

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


W_InternetCheckConnection(lpszUrl){ ;���FTP�����Ƿ������
	FLAG_ICC_FORCE_CONNECTION := 0x1
	dwReserved := 0x0
	return, DllCall("Wininet.dll\InternetCheckConnection", "Ptr", &lpszUrl, "UInt", FLAG_ICC_FORCE_CONNECTION, "UInt", dwReserved, "Int")
}

InternetFileRead( ByRef V, URL="", RB=0, bSz=1024, DLP="DLP", F=0x84000000 )
{
	SetBatchLines, -1
	Static LIB="WININET\", QRL=16, CL="00000000000000", N=""
	If ! DllCall( "GetModuleHandle", Str,"wininet.dll" )
		DllCall( "LoadLibrary", Str,"wininet.dll" )
	If ! hIO:=DllCall( LIB "InternetOpen", Str,N, UInt,4, Str,N, Str,N, UInt,0 )
		Return -1
	If ! ( ( hIU:=DllCall( LIB "InternetOpenUrl", UInt,hIO, Str,URL, Str,N, Int,0, UInt,F , UInt,0 ) ) || ErrorLevel )
		Return 0 - ( !DllCall( LIB "InternetCloseHandle", UInt,hIO ) ) - 2
	If ! ( RB )
	If ( SubStr(URL,1,4) = "ftp:" )
		CL := DllCall( LIB "FtpGetFileSize", UInt,hIU, UIntP,0 )
	Else If ! DllCall( LIB "HttpQueryInfo", UInt,hIU, Int,5, Str,CL, UIntP,QRL, UInt,0 )
		Return 0 - ( !DllCall( LIB "InternetCloseHandle", UInt,hIU ) ) - ( !DllCall( LIB "InternetCloseHandle", UInt,hIO ) ) - 4
	VarSetCapacity( V,64 ), VarSetCapacity( V,0 )
	SplitPath, URL, FN,,,, DN
	FN:=(FN ? FN : DN), CL:=(RB ? RB : CL), VarSetCapacity( V,CL,32 ), P:=&V,
	B:=(bSz>CL ? CL : bSz), TtlB:=0, LP := RB ? "Unknown" : CL, %DLP%( True,CL,FN )
	Loop
	{
		If ( DllCall( LIB "InternetReadFile", UInt,hIU, UInt,P, UInt,B, UIntP,R ) && !R )
			Break
		P:=(P+R), TtlB:=(TtlB+R), RemB:=(CL-TtlB), B:=(RemB<B ? RemB : B), %DLP%( TtlB,LP )
		Sleep -1
	}
	TtlB<>CL ? VarSetCapacity( T,TtlB ) DllCall( "RtlMoveMemory", Str,T, Str,V, UInt,TtlB ) . VarSetCapacity( V,0 ) . VarSetCapacity( V,TtlB,32 ) . DllCall( "RtlMoveMemory", Str,V , Str,T, UInt,TtlB ) . %DLP%( TtlB, TtlB ) : N
	If ( !DllCall( LIB "InternetCloseHandle", UInt,hIU ) ) + ( !DllCall( LIB "InternetCloseHandle", UInt,hIO ) )
		Return -6
	Return, VarSetCapacity(V)+((ErrorLevel:=(RB>0 && TtlB<RB)||(RB=0 && TtlB=CL) ? 0 : 1)<<64)
}

DLP(WP=0, LP=0, MSG="")
{
	global INI,FN,T1,T2,WP1,WP2,SP
	GuiControl, Text, Label1, �������أ�%FN%
	GuiControl,, MyProgress, % Round(WP/LP*100)
	DllCall("QueryPerformanceCounter", "Int64*", T2)
	DllCall("QueryPerformanceFrequency", "Int64*", TI)
	WP2:=WP
	if ((T:=(T2-T1)/TI) >=1)
	{
		SP:=Round(((WP2-WP1)/1024)/T,2)
		T1:=T2
		WP1:=WP2
	}
	WP:= ((WP:= Round(WP/1024)) < 1024) ? WP . " KB" : Round(WP/1024, 2) . " MB"
	LP:= ((LP:= Round(LP/1024)) < 1024) ? LP . " KB" : Round(LP/1024, 2) . " MB"
	GuiControl, Text, Label2, %WP% / %LP%
	GuiControl, Text, Label3, %SP% KB/S
}

VarZ_Save( byRef V, File="" ) { ; www.autohotkey.net/~Skan/wrapper/FileIO16/FileIO16.ahk
Return ( ( hFile := DllCall( "_lcreat", AStr,File, UInt,0 ) ) > 0 )
 ? DllCall( "_lwrite", UInt,hFile, Str,V, UInt,VarSetCapacity(V) )
 + ( DllCall( "_lclose", UInt,hFile ) << 64 ) : 0
}

Update(URL){
	static req := ComObjCreate("Msxml2.XMLHTTP")
	req.open("GET",URL,false)
	req.Send()
	return req.responseText
}

Gui_Submit(Access){
	Gui,PS:Submit,NoHide
	IniRead,IniList,%A_ScriptDir%\ini.ini
	global GuiArr
	Loop,Parse,IniList,`n
	{
		IniRead,Gui_Key,%A_ScriptDir%\ini.ini,%A_LoopField%,Key
		IniRead,Gui_Way,%A_ScriptDir%\ini.ini,%A_LoopField%,Way,%A_Space%
		IniRead,Gui_Label,%A_ScriptDir%\ini.ini,%A_LoopField%,Label
		if Access
		{	;д�����õ�INI ���� ��������ʱִ��
			if (Gui_Way="Hotkey")
				Hotkey,%Gui_Key%,Off
			GuiControlGet,Gui_Key,,%A_LoopField%
			if (Gui_Way="Hotkey"){
				Hotkey,%Gui_Key%,%Gui_Label%
				Hotkey,%Gui_Key%,On
			}
			IniWrite,%Gui_Key%,%A_ScriptDir%\ini.ini,%A_LoopField%,Key
		}else{ ;��ȡ���õ�GUI ���� ֻ�ڽű���ʼʱִ��
			if (Gui_Way="Hotkey"){
				Hotkey,%Gui_Key%,%Gui_Label%
				Hotkey,%Gui_Key%,On
				GuiControl,PS:,%A_LoopField%,%Gui_Key%
			}
			if (Gui_Way="DDL"){
				IniRead,Gui_var,%A_ScriptDir%\ini.ini,%A_LoopField%,var
				GuiControl,PS:,%A_LoopField%,|%Gui_var%
				GuiControl,PS:Choose,%A_LoopField%,%Gui_Key%
				StringSplit,String,Gui_var,|
				GuiArr[A_LoopField]:=String%Gui_Key%
				;MsgBox % Gui_Way "`n" Gui_Key "`n" A_LoopField "`n" Gui_var "`n" GuiArr[A_LoopField]
			}
		}
	}
	return
}

Activate(Ranking){
	global lists
	global listv
	global Hots
	global Arrays
	Activate:=listv[Ranking]
	WinActivate,ahk_id %Activate%
	loop,9
		Hotkey,%A_Index%,off
	For k,v in Arrays{
		IfWinNotExist,ahk_id %v%
			Arrays.Delete(k)
		IfWinNotExist,%k%
			Arrays.Delete(k)
	}
	Hots:=
	Cancel()
	return
}

Cancel(){
	global DDL	
	if DDL=ListView
		Gui,Cancel
	else
		TrayTip
}


GetArray(){
	Array := Object()
	TrayTip,��������׼��,,,2
	DetectHiddenWindows,Off
	WinGet,id,List,,, Program Manager
	DetectHiddenWindows,On
	loop,%id%
	{
		this_id := id%A_Index%
		WinGet,idexe,ProcessName,ahk_id %this_id%
		WinGet,WinList,List,ahk_exe %idexe%
		loop,%WinList% {
			id:=WinList%A_Index%
			WinGetTitle,Title,ahk_id %id%
			if Title
				Array[Title] := id
		}
	}
	return Array
}
