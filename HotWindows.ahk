/*�㻹��Ϊÿ��Ѱ�Ҹ���������ڶ�������
*����Ϊ��ͻ�QQ��ͨѰ�Ҵ��ڶ�������
*�� ����ʲô����� ֱ�ӿ�ʹ��˵������ɣ�
*��������� ������������ ��ӭ��ǿ��
*�����ű���ȴ���������ʾ׼�����
*����Ϊ���ô��ڱ�����ĸ��������
*���缤��AutoHotkey�߼�Ⱥ����
*��ס�ո��ڵ��GJQ���߼�Ⱥ����ƴ�������ToolTip��TrayTip��ʾһ���б�
*���û���б�˵��û���������ֵĴ���
*����б����ж������������ְ������ּ�����Ӧ�Ĵ���
*������輤��Ĵ���Ϊͷ�����ɿ��ո�󼴿ɼ���
*ע��WIN7ϵͳ������ʾΪ���� WIN10��ʾΪ��Ŀ
*WIN10�޸�PatternΪ�� WIN7�޸�PatternΪ1
*/
DetectHiddenWindows,On
#Persistent
#SingleInstance force
#UseHook
#InstallKeybdHook

Key=Tab
Layout=qwertyuiopasdfghjklzxcvbnm
Loop,Parse,Layout
	Hotkey,~%A_LoopField%,Layout

Hotkey,!q,WinMinimize
Hotkey,!w,WinMaximize
Hotkey,!e,WinMove

Pattern :=1
tcmatch := "Tcmatch.dll"
hModule := DllCall("LoadLibrary", "Str", tcmatch, "Ptr")
Arrays:=GetArray()
	;For k,v in Arrays
	;	MsgBox % k "`n" v
TrayTip,HotWindows,�����Ѿ�����׼��,,1
loop,Parse,Layout
	Hotkey,~Space & ~%A_LoopField%,HotWindows
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
}
}
return

WinMinimize: 	 ;������С��
	WinMinimize A
return
WinMaximize:	;�������
	WinMaximize A
return
Layout:
DetectHiddenWindows,off
if GetKeyState(key,"P")
	if (A_TimeSincePriorHotkey<"200") and (ThisHotkey=A_ThisHotkey){
		StringReplace,ThisHotkey,A_ThisHotkey,~
		RegRead,boss_exe,HKEY_LOCAL_MACHINE,SOFTWARE\TestKey,bossexe%ThisHotkey%
		RegRead,boss_path,HKEY_LOCAL_MACHINE,SOFTWARE\TestKey,bosspath%ThisHotkey%
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
		;TrayTip,,%A_ThisHotkey% "`n" %A_TimeSincePriorHotkey% 
	}
DetectHiddenWindows,On
return
WinMove:	;�ѵ�ǰ���ڻ�ԭ���Ƶ���Ļ���м�!
	D_Width:=(A_ScreenWidth/2)
	D_Height:=(A_ScreenHeight/2)
	WinGetPos,X,Y,Width,Height,A
	WinMove,A,,A_ScreenWidth/8,A_ScreenHeight/8,(A_ScreenWidth/8)*6,(A_ScreenHeight/8)*6
return


HotWindows:
	List :=
	marry :=
	Spacel :=
	StringRight,key,A_ThisHotkey,1
	keys = %keys%%key%
	keys := Trim(keys)
	D_VAR := StrLen(keys)
	for k,v in Arrays{
		if (matched := DllCall("Tcmatch\MatchFileW","WStr",keys,"WStr",k)){
			marry++
			list=%list%`n%marry%-%k%
		}
	if marry=9
		break
	}
	list := Trim(list,"`n")
	StringSplit,lis,List,`n
	;if (A_OSVersion="WIN_7")
		if Pattern{
			TrayTip,,%list%
		}else{
			WinGetPos,x,y,,,A
			ToolTip,%list%,%x%,%y%
		}
	if (marry="1"){
		StringTrimLeft,lismarry,lis1,2
		GetId:=Arrays[lismarry]
		WinActivate,ahk_id %GetId%
		Spacel:=1
	}
	if (D_VAR="1"){
		loop,9{
			Hotkey,%A_Index%,Table
			Hotkey,%A_Index%,On
		}
		goto,WaitL
	}
return

WaitL:
	KeyWait,Space,L
	if not Spacel{
		StringTrimLeft,liswait,lis1,2
		GetId:=Arrays[liswait]
		WinActivate,ahk_id %GetId%
	}
	if D_VAR>=2
		loop,9
			Hotkey,%A_Index%,off
	ToolTip
	TrayTip
	keys:=
return
Table:
	StringTrimLeft,liss,lis%A_ThisHotkey%,2
	GetId:=Arrays[liss]
	WinActivate,ahk_id %GetId%
	loop,9
		Hotkey,%A_Index%,off
	Spacel:=1
return





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
