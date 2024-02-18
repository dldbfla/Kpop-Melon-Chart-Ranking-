#NoEnv
#SingleInstance force
SetBatchLines, -1
ListLines, Off
#KeyHistory 0
Process, Priority,, High
;---------------------------------------------------------------

Menu, Tray, Icon, %A_ScriptDir%\AutoHotKey Light.png
Gui, Font, S10
Gui, Add, Button, x910 y10 w100 h50 gSave, ��Ʈ ����
Gui, Add, Button, x800 y10 w100 h50 gClipBoard, Ŭ������ ����
Gui, Add, Text, x250 y750 w600 h20 +Center, �ش� ��Ʈ�� ��� Ȩ���������� ������ ���̸�`, 1�ð� ������ ������Ʈ �˴ϴ�.
Gui, Font, W700
Gui, Add, Button, x1020 y10 w100 h50 gUpdate, ��Ʈ ������Ʈ
Gui, Add, ListView, x22 y70 w1100 h670 ReadOnly Grid , ����|����|���|����|�ٹ���
Gui, Font, S20
Gui, Add, Text, x22 y38 w350 h30 c13C7A3 vTime ,
Gui, Add, Picture, x+10 y38 h-1 gLogo, %A_ScriptDir%\melon.png
Gui, Show, h780 w1140, ��� �ǽð� TOP 100
chart()

Return


Logo:
if(A_GuiEvent = "DoubleClick")
{
	run, https://www.melon.com/chart/
}
return

Update:
{
	chart()
}
return

chart()
{
	LV_Delete()
	SplashTextOn, , , ��� ������Ʈ ��...

	Melon := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	Melon.Open("get", "http://www.melon.com/chart/index.htm")
	Melon.Send("")
	winHttp.WaitForResponse( )
	RegExMatch(Melon.responseText, "<span class=\Cyear\C>(.*?)</span>", Year)
	RegExMatch(Melon.responseText, "<span class=\Chour\C>(.*?)</span>", Hour)
	GuiControl, , Time, %Year1% %Hour1%�� ����

	p := 1, m := ""
	while p:= RegExMatch(Melon.responseText, "<span\sclass=\Crank\s\C>(\d+)</span><span.+>", m, p + StrLen(m))
	{
		LV_Add( , m1)
		LV_Modify(A_Index, , , m1)
	}

	p := 1, m := ""
	While, p:= RegExMatch(Melon.responseText, "<span\stitle=\C(.*?)\C\sclass=\Crank_wrap\C>", m, p + StrLen(m))
	{
		LV_Modify(A_Index, , ,m1)
	}

	p := 1, m := ""
	While, p:= RegExMatch(Melon.responseText, "<div\sclass=\Cellipsis\srank01\C><span>\s+<a\shref.+\stitle=\C.+\C>(.*?)</a>", m, p + StrLen(m))
	{
		LV_Modify(A_Index, , , ,m1)
	}

	p := 1, m := ""
	While, p:= RegExMatch(Melon.responseText, "<div\sclass=\Cellipsis\srank02\C>\s+<a\shref.+\stitle=\C.+\C>(.*?)</a>", m, p + StrLen(m))
	{
		LV_Modify(A_Index, , , , ,m1)
	}

	p := 1, m := ""
	While, p:= RegExMatch(Melon.responseText, "<div\sclass=\Cellipsis\srank03\C>\s+<a\shref.+\stitle=\C.+\C>(.*?)</a>", m, p + StrLen(m))
	{
		LV_Modify(A_Index, , , , , ,m1)
	}
	LV_ModifyCol()
	LV_ModifyCol(1, "40 right")
	LV_ModifyCol(2, "90 right")
	SplashTextOff
}
return

ClipBoard:
{
	Raw_Count := LV_GetCount()
	if(Raw_Count = 0)
	{
		MsgBox, , ����!, ���� ��Ʈ�� �ҷ����� �ʾҽ��ϴ�.
		return
	}
	Loop, %Raw_Count%
	{
		LV_GetText(Save_Num, A_Index, 1), LV_GetText(Save_Rank, A_Index, 2), LV_GetText(Save_SongName, A_Index, 3), LV_GetText(Save_Artist, A_Index, 4), LV_GetText(Save_Album, A_Index, 5)

		Save_Text .= Save_Num "`t" Save_Rank "`t" Save_SongName "`t" Save_Artist "`t" Save_Album "`n"
	}
	Clipboard := Save_Text
	MsgBox, , ���� �Ϸ�, ��� Top 100 ��Ʈ�� Ŭ������� ����Ǿ����ϴ�.
}
return

Save:
{
	Raw_Count := LV_GetCount()
	if(Raw_Count = 0)
	{
		MsgBox, , ����!, ���� ��Ʈ�� �ҷ����� �ʾҽ��ϴ�.
		return
	}
	IfExist, Melon Top 100.cvs
	{
		FileDelete, Melon Top 100.cvs
	}
	Save_Text := "����" "," "����" "," "���" "," "����" "," "�ٹ���" "`n"
	Loop, %Raw_Count%
	{
		LV_GetText(Save_Num, A_Index, 1), LV_GetText(Save_Rank, A_Index, 2), LV_GetText(Save_SongName, A_Index, 3), LV_GetText(Save_Artist, A_Index, 4), LV_GetText(Save_Album, A_Index, 5)
		Save_Text .= Save_Num "," Save_Rank "," Save_SongName "," Save_Artist "," Save_Album "`n"
	}
	FileAppend, %Save_Text%, Melon Top 100.cvs, UTF-8
}
return

ExitMenuHandler:
GuiClose:
{
	ExitApp
}
return

;[��ó] ��� �ǽð� TOP 100 ��Ʈ - ��¯����
;[��ũ] http://apz.kr/mybbs/648243