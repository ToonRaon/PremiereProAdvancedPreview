#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#SingleInstance force
#MaxHotkeysPerInterval 1000


VERSION = 1.0
optionFileDir = %A_MyDocuments%\프리미어_프로_프리뷰_개선_환경설정.txt

if(!FileExist(optionFileDir))  ; 옵션 파일 없으면 생성 (최초 실행시)
{
	createOptionFile(optionFileDir)
}


stage := 0
flag := false
OnOffCheckBox := getOption(optionFileDir, "Enabled", 1)
LeftTopX := getOption(optionFileDir, "LeftTopX", -1)
LeftTopY := getOption(optionFileDir, "LeftTopY", -1)
RightBottomX :=  getOption(optionFileDir, "RightBottomX", -1)
RightBottomY := getOption(optionFileDir, "RightBottomY", -1)
isFocusOn := getOption(optionFileDir, "Focus", 0)

guiWidth := 360
guiHeight := 210

notCompleted = 좌표 설정을 모두 완료하여 주세요.
completed = 프리미어 프리뷰 개선이 적용된 상태입니다.



; 트레이 우클릭 메뉴 설정
Menu, Tray, Add
Menu, Tray, Add, 설정, ShowGUI
Menu, Tray, Add, 사용법, LaunchGuide
Menu, Tray, Add, 현재 버전: %VERSION%, NONE


goto InitGui ; Gui 초기화

createOptionFile(optionFileDir)
{
	FileAppend, Enabled:1`n, %optionFileDir%
	FileAppend, Focus:0`n, %optionFileDir%
	FileAppend, LeftTopX:-1`n, %optionFileDir%
	FileAppend, LeftTopY:-1`n, %optionFileDir%
	FileAppend, RightBottomX:-1`n, %optionFileDir%
	FileAppend, RightBottomY:-1`n, %optionFileDir%
}

getOption(optionFileDir, key, default)
{
	Loop
	{
		FileReadLine, line, %optionFileDir%, %A_Index%
		if(ErrorLevel)
		{
			break
		}
		IfInString, line, %key%
		{
			return StrSplit(line, ":")[2]
		}
	}
	
	return default
}

setOption(optionFileDir, key, value)
{
	tempFile = %A_MyDocuments%\tempDuplicatedOptionFile.txt
	FileCopy, %optionFileDir%, %tempFile%
	FileDelete, %optionFileDir%
	Loop
	{
		FileReadLine, line, %tempFile%, %A_Index%
		if(ErrorLevel)
		{
			break
		}
		IfInString, line, %key%
		{
			FileAppend, %key%:%value%`n, %optionFileDir%
		}
		else
		{
			FileAppend, %line%`n, %optionFileDir%
		}
	}
	FileDelete, %tempFile%
}

isEnglishNow()
{
	if(IME_CHECK("A") == 0)
	{
		return true
	}
	else
	{
		return false
	}
}

changeLanguage() ; 한영 전환
{
	Send {vk15sc138}
}

; 한영 전환 스크립트출처: https://blog.hangyeong.com/821
IME_CHECK(WinTitle) ; 한영이 영어 상태면 "0" return 
{
	WinGet,hWnd,ID,%WinTitle%
	Return Send_ImeControl(ImmGetDefaultIMEWnd(hWnd),0x005,"")
}

Send_ImeControl(DefaultIMEWnd, wParam, lParam)
{
	DetectSave := A_DetectHiddenWindows
	DetectHiddenWindows,ON
	SendMessage 0x283, wParam,lParam,,ahk_id %DefaultIMEWnd%
	if (DetectSave <> A_DetectHiddenWindows)
		DetectHiddenWindows,%DetectSave%
	return ErrorLevel
}

ImmGetDefaultIMEWnd(hWnd)
{
	return DllCall("imm32\ImmGetDefaultIMEWnd", Uint,hWnd, Uint)
}

NONE:
return

InitGui:
guiPadding := 20

Gui, Color, FAFAFA
Gui, Margin, %guiPadding%, %guiPadding%

; Gui, Add, CheckBox, vOnOffCheckBox gSubmitGui Checked, 활성화 여부
Gui, Add, Button, gLaunchGuide w315, 사용법 (유튜브 영상)
Gui, Add, Checkbox, vFocusCheckBox gSubmitGui, 스크롤시 프리뷰창 강제 포커싱 여부
if(isFocusOn)
{
	GuiControl, , FocusCheckBox, 1
}
Gui, Add, Text, , 프리뷰 왼쪽 위 좌표
Gui, Add, Edit, w50 x+5 ReadOnly vLeftTopX, %LeftTopX%
Gui, Add, Text, x+5, `,
Gui, Add, Edit, w50 x+5 ReadOnly vLeftTopY, %LeftTopY%
Gui, Add, Button, x+5 gSetupLeftTopCoordinate, 좌표 설정
Gui, Add, Text, xs, 프리뷰 오른쪽 아래 좌표
Gui, Add, Edit, w50 x+5 ReadOnly vRightBottomX, %RightBottomX%
Gui, Add, Text, x+5, `,
Gui, Add, Edit, w50 x+5 ReadOnly vRightBottomY, %RightBottomY%
Gui, Add, Button, x+5 gSetupRightBottomCoordinate, 좌표 설정
Gui, Add, Text, xs w300 cRed vCompleteText, ...
if(LeftTopX == -1 or RightBottomX == -1)
{
	GuiControl, , CompleteText, %notCompleted%
}
else
{
	GuiControl, , CompleteText, %completed%
}

Gui, Submit, NoHide
return

LaunchGuide:
Run https://youtu.be/T1nRD5BaVyE
return

SubmitGui:
if(A_GuiControl == "FocusCheckBox")
{
	GuiControlGet, checked, , FocusCheckBox
	isFocusOn := checked
	setOption(optionFileDir, "Focus", isFocusOn)
}
Gui, Submit, NoHide
return

SetupLeftTopCoordinate:
Gui, Minimize
stage := 1
MsgBox, 프리미어 프로 프리뷰 창의 왼쪽 위에 마우스를 가져다 댄 후 키보드의 'Pause'을 누르세요.
WinActivate, ahk_exe Adobe Premiere Pro.exe
return

SetupRightBottomCoordinate:
Gui, Minimize
stage := 2
MsgBox, 프리미어 프로 프리뷰 창의 오른쪽 아래에 마우스를 가져다 댄 후 키보드의 'Pause'을 누르세요.
WinActivate, ahk_exe Adobe Premiere Pro.exe
return

ShowGUI:
Gui, Show, w%guiWidth% h%guiHeight%, 프리미어 프로 프리뷰 개선
return

#If WinActive("ahk_exe Adobe Premiere Pro.exe")

isMouseOverOnPreviewWindow(x1, y1, x2, y2) {
	MouseGetPos, x, y
	
	if(x1 <= x and x <= x2 and y1 <= y and y <= y2) {
		return true
	} else {
		return false
	}
}

Pause::
Switch stage
{
	Case 1:
		MouseGetPos, x, y
		LeftTopX := x
		LeftTopY := y
		stage := 0
		setOption(optionFileDir, "LeftTopX", x)
		setOption(optionFileDir,  "LeftTopY", y)
		GuiControl, , LeftTopX, %x%
		GuiControl, , LeftTopY, %y%
		if(RightBottomX != -1)
		{
			GuiControl, , CompleteText, %completed%
		}
		Gui, Restore
	Case 2:
		MouseGetPos, x, y
		RightBottomX := x
		RightBottomY := y
		stage := 0
		setOption(optionFileDir, "RightBottomX", x)
		setOption(optionFileDir,  "RightBottomY", y)
		GuiControl, , RightBottomX, %x%
		GuiControl, , RightBottomY, %y%
		if(LeftTopX != -1)
		{
			GuiControl, , CompleteText, %completed%
		}
		Gui, Restore
	Default:
		goto, ShowGUI
}
return

#if WinActive("ahk_exe Adobe Premiere Pro.exe") and isMouseOverOnPreviewWindow(LeftTopX, LeftTopY, RightBottomX, RightBottomY)
WheelUp::
if(isFocusOn)
{
	focusOnPreview(RightBottomX, RightBottomY)
}
Send {NumPadAdd}
return

WheelDown::
if(isFocusOn)
{
	focusOnPreview(RightBottomX, RightBottomY)
}
Send {NumPadSub}
return

focusOnPreview(rbX, rbY)
{
	MouseGetPos, x, y
	MouseClick, Left, rbX, rbY, , 0
	MouseMove, x, y, 0
}

MButton::
flag = true
eng := isEnglishNow()
if(!eng)
{
	changeLanguage()
}
Send h
if(!eng)
{
	changeLanguage()
}
MouseClick, Left, , , , , D
return

#if WinActive("ahk_exe Adobe Premiere Pro.exe") and flag
MButton Up::
MouseClick, Left, , , , , U
eng := isEnglishNow()
if(!eng)
{
	changeLanguage()
}
Send v
if(!eng)
{
	changeLanguage()
}
flag = false
return