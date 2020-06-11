; This script uses Path of Exile Weapon DPS Calculator 
; (https://pdejan.github.io/poe_wdps/)
; This script uses Aeons PoEDisplay.ahk
; https://gist.github.com/aeons/7432713

#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
#Persistent ; Stay open in background
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
StringCaseSense, On ; Match strings with case.

MouseMoveThreshold := 5 ; Pixels must move to auto-dismiss tooltip
Menu, tray, Tip, PoE DPS Display (by dilof and Mih0l) ; Menu tooltip
	
#Include %A_ScriptDir%\data\GemQualityList.txt ;forked and edited from PoE_ItemInfo

FindItemStats(wb)
{
    IsWeapon := False
	IsGem =: False
    PhysLo := 0
    PhysHi := 0
	ChaosLo :=0
	ChaosHi :=0
    Quality := 0
    AttackSpeed := 0
	WebNav := 0
	
	Global QualityList
	
	Loop, Parse, Clipboard, `n, `r
	{

		; Clipboard must have "Rarity:" in the first line
		If A_Index = 1
		{
			IfNotInString, A_LoopField, Rarity:
			{
				Exit
			}
			Else
			{
				IfInString, A_LoopField, Gem
				{
					IsGem := True
				}
			}
		}
		
		; Get Gem name
		If (A_Index = 2 && IsGem = true)
		{
			GemName := A_LoopField
		}
		
		; Get item influence Elder
		IfInString, A_LoopField, Elder Item
		{
			itemInfluence = Elder Item
			Continue
		}	
		
		; Get item influence Shaper
		IfInString, A_LoopField, Shaper Item
		{
			itemInfluence = Shaper Item
			Continue
		}	
		
		; Get item influence Redeemer
		IfInString, A_LoopField, Redeemer Item
		{
			itemInfluence = Redeemer Item
			Continue
		}	
		
		; Get item influence Warlord
		IfInString, A_LoopField, Warlord Item
		{
			itemInfluence = Warlord Item
			Continue
		}	
		
		; Get item influence Crusader
		IfInString, A_LoopField, Crusader Item
		{
			itemInfluence = Crusader Item
			Continue
		}	
		
		; Get item influence Hunter
		IfInString, A_LoopField, Hunter Item
		{
			itemInfluence = Hunter Item
			Continue
		}	
		
		; Get item influence Synthesised
		IfInString, A_LoopField, Synthesised Item
		{
			itemInfluence = Synthesised Item
			Continue
		}	

        ; Get quality
        IfInString, A_LoopField, Quality:
        {
            StringSplit, Arr,  A_LoopField, %A_Space%, +`% 
            Quality := Arr2
            Continue
        }
 
        ; Get total chaos damage
        IfInString, A_LoopField, Chaos Damage:
		{
            IsWeapon = True
            StringSplit, Arr, A_LoopField, %A_Space%
            StringSplit, Arr, Arr3, -
            ChaosLo := Arr1
            ChaosHi := Arr2
            Continue
        }
 
        ; Get total physical damage
        IfInString, A_LoopField, Physical Damage:
        {
            IsWeapon = True
            StringSplit, Arr, A_LoopField, %A_Space%
            StringSplit, Arr, Arr3, -
            PhysLo := Arr1
            PhysHi := Arr2
            Continue
        }
		
        ; These only make sense for weapons
        If IsWeapon 
        {
            ; Get attack speed
            IfInString, A_LoopField, Attacks per Second:
            {
                StringSplit, Arr, A_LoopField, %A_Space%
                AttackSpeed := Arr4
                Continue
            }
		}
	}

	global X
	global Y
	MouseGetPos, X, Y

	wb := ComObjCreate("InternetExplorer.Application")  ;// Create an IE object
	if WebNav = 0
	{
	wb.Visible := False                                ;// Make the IE object visible
	wb.Navigate("https://pdejan.github.io/poe_wdps/")                   ;// Navigate to a webpage
	while wb.busy or wb.ReadyState != 4
	sleep 10
	WebNav := 1
	}
	wb.document.querySelector("textarea").value := Clipboard
	wb.document.getElementsByClassName("btn")[0].click()
		
	phydps := wb.document.getElementById("phydps").innerText
	eledps := wb.document.getElementById("eledps").innerText
	totaldps := wb.document.getElementById("totaldps").innerText	
	chaosdps := ((ChaosLo + ChaosHi ) / 2) * AttackSpeed	
	phydps := format("{:0.2f}", phydps)
	eledps := format("{:0.2f}", eledps)
	totaldps := format("{:0.2f}", totaldps)
	chaosdps := format("{:0.2f}", chaosdps)
	
	;Q20 calculations
	if (Quality < 20) 
	{
		PhysBaseCurrent := ((PhysLo + PhysHi) / 2)
		PhysBase := PhysBaseCurrent * 100 / (100 + Quality)
		Q20Phys := PhysBase * 1.2
		Q20PhysLo := Q20Phys - (PhysBase * 0.25)
		Q20PhysLo := format("{:0.1f}", Q20PhysLo)
		Q20PhysHi := Q20Phys + (PhysBase * 0.25)
		Q20PhysHi := format("{:0.1f}", Q20PhysHi)
		StringRight, Q20PhysLoEnds, Q20PhysLo, 1
		StringRight, Q20PhysHiEnds, Q20PhysHi, 1
		AddEnds := Q20PhysLoEnds + Q20PhysHiEnds
			if (AddEnds >= 10) 
			{
				Q20PhysLo := ceil(Q20PhysLo)
				Q20PhysHigh := ceil(Q20PhysHigh)
			}
		ClipboardQ20 := Clipboard
		Loop, Parse, ClipboardQ20, `n, `r
		{
			IfInString, A_LoopField, Quality:
			{
				StringReplace, ClipboardQ20, ClipboardQ20, %Quality%, 20
			}	
			IfInString, A_LoopField, Physical Damage:
			{
				StringReplace, ClipboardQ20, ClipboardQ20, %PhysLo%, %Q20PhysLo%
				StringReplace, ClipboardQ20, ClipboardQ20, %PhysHi%, %Q20PhysHi%
			}	
		}
		wb.document.querySelector("textarea").value := ClipboardQ20
		wb.document.getElementsByClassName("btn")[0].click()
	
		Q20phydps := wb.document.getElementById("phydps").innerText
		Q20totaldps := wb.document.getElementById("totaldps").innerText	
		Q20phydps := format("{:0.2f}", Q20phydps)
		Q20totaldps := format("{:0.2f}", Q20totaldps)
		
	}
	
	if (phydps < 0) || (eledps < 0) || (totaldps < 0) 
	{
		phydps := 0
		eledps := 0
		totaldps := 0
	}	
	
	wb.quit

		if IsWeapon ; Show tooltip only if it's a weapon
		{
			if (itemInfluence = null) ; influence = 0 / dont show influence
			{
				if (Quality < 20) ; Quality < 20 - show quality DPS
				{
					TT = Phys DPS: %phydps%`nElem DPS: %eledps%`nChaos DPS: %chaosdps%`nTotal DPS: %totaldps%`nQ20 DPS: %Q20totaldps%
				}
				else 
				{
					TT = Phys DPS: %phydps%`nElem DPS: %eledps%`nChaos DPS: %chaosdps%`nTotal DPS: %totaldps%
				}
			}
			else ; If item has influence
			{
				if (Quality < 20) 
				{
					TT = Phys DPS: %phydps%`nElem DPS: %eledps%`nChaos DPS: %chaosdps%`nTotal DPS: %totaldps%`nQ20 DPS: %Q20totaldps%`n%itemInfluence%
				}
				else 
				{
					TT = Phys DPS: %phydps%`nElem DPS: %eledps%`nChaos DPS: %chaosdps%`nTotal DPS: %totaldps%`n%itemInfluence%
				}
			}
		}
		if IsGem ; Show tooltip only if it's a gem
		{
			if (QualityList[GemName] != "")
			{
				GemQualityDescription := QualityList[GemName]
			}
				TT := TT . GemQualityDescription
			}
		else ; Anything else other than weapon
		{
			return ; Don't display
		}

customTooltip(TT)
SetTimer, ToolTipTimer, 100
}

customTooltip(TT)
{
	Gui, +AlwaysOnTop -Border -SysMenu +Owner -Caption +ToolWindow
	Gui, Color, 111111
    lines := StrSplit(TT, "`n")
    Y := 8
    Loop % lines.length()
    {
        line++
        currentLine = % lines[A_Index]
        If InStr(currentLine, "Total DPS:") || InStr(currentLine, "Elder Item") || InStr(currentLine, "Shaper Item") || InStr(currentLine, "Redeemer Item") || InStr(currentLine, "Warlord Item") || InStr(currentLine, "Hunter Item") || InStr(currentLine, "Crusader Item") || InStr(currentLine, "Synthesised Item") || InStr(currentLine, "Quality") || InStr(currentLine, "Elem Dps:") || InStr(currentLine, "Chaos DPS:") || InStr(currentLine, "Q20 DPS:") ||
        {
            If InStr(currentLine, "Elder Item")
            {
				Gui, Font, s11 , Comic Sans MS
                Gui, Add, text, x8 c591f82, %currentLine%
            }
				else If InStr(currentLine, "Redeemer Item")
				{
					Gui, Font, s11 , Comic Sans MS
					Gui, Add, text, x8 c336291, %currentLine%
				}
				else If InStr(currentLine, "Warlord Item")
				{
					Gui, Font, s11 , Comic Sans MS
					Gui, Add, text, x8 c8f5004, %currentLine%
				}
				else If InStr(currentLine, "Hunter Item")
				{
					Gui, Font, s11 , Comic Sans MS
					Gui, Add, text, x8 c365c1d, %currentLine%
				}
				else If InStr(currentLine, "Crusader Item")
				{
					Gui, Font, s11 , Comic Sans MS
					Gui, Add, text, x8 c349191, %currentLine%
				}
				else If InStr(currentLine, "Synthesised Item")
				{
					Gui, Font, s11 , Comic Sans MS
					Gui, Add, text, x8 c5d8194, %currentLine%
				}
				else If InStr(currentLine, "Shaper Item")
				{
					Gui, Font, s11 , Comic Sans MS
					Gui, Add, text, x8 c30455c, %currentLine%
				}
				else If InStr(currentLine, "Total DPS:")
				{
					Gui, Add, text, x8 cf50505, %currentLine%
				}
				else If InStr(currentLine, "Elem DPS:")
				{
					Gui, Add, text, x8 ce9f505, %currentLine%
				}				
				else If InStr(currentLine, "Chaos DPS:")
				{
					Gui, Add, text, x8 c24e0da, %currentLine%
				}
				else If InStr(currentLine, "Q20 DPS:")
				{
					Gui, Add, text, x8 cb8b8b8, %currentLine%
				}
				else If InStr(currentLine, "Quality")
				{
					Gui, Font, s11 , Franklin Gothic Medium
					Gui, Add, text, x8 c1d9296, %currentLine%
				}
        }
        else
        {
			Gui, Font, s10 , Arial
            Gui, Add, text, x8 cWhite, %currentLine%
        }
        Y := Y + 15
    }
    MouseGetPos, X, Y
    X := X + 25 ; X tooltip position from mouse
    Y := Y + 25 ; Y tooltip position from mouse
	Gui, Show, NoActivate x%X% y%Y%, tooltipToggle
	WinSet, Transparent, 230, ahk_class AutoHotkeyGUI ; Set window transparency
}
 
; Tick every 100 ms
; Remove tooltip if mouse is moved
ToolTipTimer:
ToolTipTimeout += 1
MouseGetPos, CurrX, CurrY
MouseMoved := (CurrX - X)**2 + (CurrY - Y)**2 > MouseMoveThreshold**2
If (MouseMoved or ToolTipTimeout >= 100) ; 10 seconds to remove the tooltip if mouse hasn't been moved
{
    SetTimer, ToolTipTimer, Off
    ToolTipTimeout := 0
    Gui, Destroy
}
return

OnClipBoardChange:
FindItemStats(wb)