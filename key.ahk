DetectHiddenWindows,On
#Persistent
#SingleInstance force
#UseHook
#InstallKeybdHook
Menu,Tray,Click,1
Menu,Tray,Add,Hot-Windows,Menu_show
Menu,Tray,Add,��������,Auto
Menu,Tray,Add,�����ű�,Reload
Menu,Tray,Add,�˳��ű�,ExitApp
Menu,Tray,Default,Hot-Windows
Menu,Tray,NoStandard

Gui,+HwndMyGuiHwnd -MaximizeBox
Gui,Add,Text,,������
Gui,Add,DDL,vDDL1 Choose1,Tab|Space
Gui,Add,Text,,�ȼ���
Gui,Add,DDL,vDDL2 Choose1,Space|Tab
Gui,Add,Text,,��С���ȼ�
Gui,Add,Hotkey,,!q
Gui,Add,Text,,����ȼ�
Gui,Add,Hotkey,,!w
Gui,Add,Text,,�����ȼ�
Gui,Add,Hotkey,,!e
Gui,Add,Button,gSubmit w135,��������
Gui,Show

RegRead,HotRun,HKEY_CURRENT_USER,Software\Microsoft\Windows\CurrentVersion\Run,HotRun
IfExist,%HotRun%
{
	RegWrite,REG_SZ,HKEY_CURRENT_USER,Software\Microsoft\Windows\CurrentVersion\Run,HotRun,%A_ScriptFullPath%
	Menu,Tray,ToggleCheck,��������
}
;������׼��
tcmatch := "Tcmatch.dll"
hModule := DllCall("LoadLibrary", "Str", tcmatch, "Ptr")
Arrays:=GetArray()
;ע���ȼ�
Key=Tab
Hot=Space
Win=!
Layout=qwertyuiopasdfghjklzxcvbnm
Loop,Parse,Layout
	Hotkey,~%A_LoopField%,Layout
Hotkey,%Win%q,WinMinimize
Hotkey,%Win%w,WinMaximize
Hotkey,%Win%e,WinMove
TrayTip,HotWindows,�����Ѿ�����׼��`n�������ͼ������,1
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
	if (A_TimeSincePriorHotkey<"200") and (ThisHotkey=A_ThisHotkey){
		StringReplace,ThisHotkey,A_ThisHotkey,~
		RegRead,boss_exe,HKEY_LOCAL_MACHINE,SOFTWARE\TestKey,bossexe%ThisHotkey%
		RegRead,boss_path,HKEY_LOCAL_MACHINE,SOFTWARE\TestKey,bosspath%ThisHotkey%
		DetectHiddenWindows,Off
		if boss_exe
		{
			IfWinActive,ahk_exe %boss_exe%
			{
				WinMinimize
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
			TrayTip,,����%boss_exe%`n�ȼ���%A_ThisHotkey%
			WinMinimize A
		}
		DetectHiddenWindows,On
	}
}
if GetKeyState(Hot,"P"){
	StringReplace,ThisHotkey,A_ThisHotkey,~
	Hots = %Hots%%ThisHotkey%
	vars := StrLen(Hots)
	list :=
	marry :=
	if (vars="1"){
		lists := Object()
		loop,9{
			Hotkey,%A_Index%,Table
			Hotkey,%A_Index%,On
		}
		goto,WaitHot
	}
	for k,v in Arrays{
		if (matched := DllCall("Tcmatch\MatchFileW","WStr",Hots,"WStr",k)){
			marry++
			if marry>9
				list=%list%`n%k%
			else
				list=%list%`n%marry%-%k%
			lists[marry]:=v
		}
	}
	if marry{
		list := Trim(list,"`n")
		TrayTip,,%list% ; `n %A_ThisHotkey% `n %hots% `n %vars% `n %marry%
		if (marry="1")
			Activate("1")
	}
}
	ThisHotkey:=A_ThisHotkey
return

WaitHot:
	KeyWait,%Hot%,L
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
if (DDL1=DDL1){
	TrayTip,HotWindows,���������ȼ����ȼ�������ͬ,,3
	return
}else{
	Key:=DDL1
	Hot:=DDL2

}
return

Menu_show:
DetectHiddenWindows,Off
IfWinNotExist,ahk_id %MyGuiHwnd%
	Gui,Show
else
	Gui,Cancel
DetectHiddenWindows,On
return

Activate(Ranking){
	global lists
	global Hots
	global Arrays
	Activate:=lists[Ranking]
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
	TrayTip
	return
}

GetArray(){
	Array := Object()
	d := "`n"
	s := 4096  ; ���������Ĵ�С (4 KB)
	Process, Exist
	h := DllCall("OpenProcess", "UInt", 0x0400, "Int", false, "UInt", ErrorLevel, "Ptr")
	DllCall("Advapi32.dll\OpenProcessToken", "Ptr", h, "UInt", 32, "PtrP", t)
	VarSetCapacity(ti, 16, 0)
	NumPut(1, ti, 0, "UInt")
	DllCall("Advapi32.dll\LookupPrivilegeValue", "Ptr", 0, "Str", "SeDebugPrivilege", "Int64P", luid)
	NumPut(luid, ti, 4, "Int64")
	NumPut(2, ti, 12, "UInt")
	r := DllCall("Advapi32.dll\AdjustTokenPrivileges", "Ptr", t, "Int", false, "Ptr", &ti, "UInt", 0, "Ptr", 0, "Ptr", 0)
	DllCall("CloseHandle", "Ptr", t)
	DllCall("CloseHandle", "Ptr", h)
	hModule := DllCall("LoadLibrary", "Str", "Psapi.dll")  ; ͨ��Ԥ��������������
	s := VarSetCapacity(a, s)
	c := 0
	DllCall("Psapi.dll\EnumProcesses", "Ptr", &a, "UInt", s, "UIntP", r)
	loop, % r // 4
	{
		id := NumGet(a, A_Index * 4, "UInt")
		h := DllCall("OpenProcess", "UInt", 0x0010 | 0x0400, "Int", false, "UInt", id, "Ptr")
		if !h
			continue
		VarSetCapacity(n, s, 0)
		e := DllCall("Psapi.dll\GetModuleBaseName", "Ptr", h, "Ptr", 0, "Str", n, "UInt", A_IsUnicode ? s//2 : s)
		if !e
			if e := DllCall("Psapi.dll\GetProcessImageFileName", "Ptr", h, "Str", n, "UInt", A_IsUnicode ? s//2 : s)
				SplitPath n, n
		DllCall("CloseHandle", "Ptr", h)  ; �رս��̾���Խ�Լ�ڴ�
		if (n && e)  ; ���ӳ���ǿյ�, ����ӵ��б�:
		{
			l .= n . d, c++
		}
	}
	DllCall("FreeLibrary", "Ptr", hModule)  ; ж�ؿ����ͷ��ڴ�
	Sort,l,U
	even:=100/(c-ErrorLevel)
	evens:=0
	loop, Parse,l,`n
	{
		WinGet,WinList,List,ahk_exe %A_LoopField%
		loop,%WinList% {
			id:=WinList%A_Index%
			WinGetTitle,Title,ahk_id %id%
			if Title
				Array[Title] := id
		}
		evens:=evens+even
		event:=Ceil(evens)
		TrayTip,��������׼��,%event%,,2
	}
	return Array
}
