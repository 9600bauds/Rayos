#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance force ; Close old versions of this script automatically.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2 ; Match window titles anywhere, not just at the start.

#Include modules\searchTypes.ahk ; Defines search types

#include modules\ventanas\ModificarArticulo.ahk
#include modules\ventanas\Proovedores.ahk

;{ Globals (most of these are effectively defines)

global vReporteArticulos_id := "Artículos de :"
global vReporteArticulos_proovedoresHabituales := "Button18"

global vReporteArticulos_planilla := "TXBROWSE1"
global vReporteArticulos_modificar = "TBTNBMP57"

global vNotepad_id := "ahk_class Notepad"
global vWord_id := " - Word"
global vCalc_id := "OpenOffice Calc"
global vCalc_buscar := "Find & Replace"
global vCalc_main := "ahk_class SALFRAME" ;Precisamente la planilla principal, no ningún diálogo
global vAdobe_id := "Adobe Acrobat"
global vAdobeBuscar_id := "Buscar ahk_exe Acrobat.exe"
global vAdobeBuscar_ok := "Button18"
global vAdobeBuscar_input := "Edit5"
global vAdobeBuscar_resultados := "Static12"

global vFacturaProov_id := "FACTURA  Proveedor.Nueva" ;sic
global vFacturaProovNuevo_id := "Nuevo"
global vFacturaProovModif_id := "Modificación"

global suppressWarnings := false
global autoPilot := false
global forceSeek := false
global overrideMiddleClick := true

global modificadoresText := "+0%" ;These two should be equivalent and are only set once in SetModificadores().
global modificadoresMult := 1

global lastPercent := 0
global preciosGuardados := {}

global PostSearchString := ""

global working := false
global shouldStop := false

global lastSeekCol := ""
;}

;{ Modificadores de Precio
SetModificadores(modificadoresInput := "", displayMessage := true){
	if(modificadoresInput == "")
	{
		explanation := "Ingrese una lista de aumentos/recargos.`nEjemplos de lista de aumentos/recargos válidos:`n+15%`n-20% +10%`n-16.66+15"
		InputBox, modificadoresInput, Descuento Básico, %explanation%,,,,,,,,%modificadoresText%
		if ErrorLevel
			return ;Cancel
	}
	
	tempTally := 1
	tempStr := ""
	tempPercent := 0
	for index, match in AllRegexMatches(modificadoresInput, "[+-]+\d+\.?\d*")
	{
		tempParsed := Percent2Multiplier(match)
		if(not tempParsed)
		{
			MsgBox, SetModificadores - Multiplicador inválido. (%match%)
			return -1
		}
		tempTally := tempTally * tempParsed
		tempStr = %tempStr% %match%`%
	}
	if(tempTally <= 0)
	{
		MsgBox, SetModificadores - Modificadores inválidos. (multiplicador resultante: %tempTally%)
		return -1
	}
	
	
	relativeDiff := Round((tempTally - modificadoresMult) / modificadoresMult * 100, 2)
	modificadoresText := tempStr
	modificadoresMult := tempTally
	tempPercent := Multiplier2Percent(modificadoresMult)
	RegWrite, REG_SZ, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedModificadores, %modificadoresText%
	if(displayMessage)
	{
		MsgBox, Modificadores actualizados. Nuevos modificadores:`n%modificadoresText%`n(%tempPercent% / x%modificadoresMult%)`nDiferencia relativa: %relativeDiff%`%
	}
}

Percent2Multiplier(percent){
	percent := RegExReplace(percent, "[^0-9|\-|.]") ;Sólo numeros.
	return (100+percent)/100
}
Multiplier2Percent(multiplier){
	multiplier := multiplier * 100
	return % multiplier . "%"
}	
;}

;{ Alias
ParseAlias(alias){
	alias := RegExReplace(alias, "^0") ;Remove leading zero.
	alias := RegExReplace(alias, "[ \t]+$") ;Remove trailing whitespace.	
	
	;if(InStr(alias, "NO TRAER") or InStr(alias, "NO COMPRAR") or RegExMatch(alias, "^[-]+$")){
	;	return "NEXT!"
	;}
	
	alias = % executeSearchTypes(alias)
	
	return alias
}

GetAlias(parseAfter := true, checkNota := true){
	if(shouldStop())
	{
		return
	}
		
	aliasText := ""
	if(WinExist(vReporteArticulos_id))
	{
		if(not WinExist(vModifArticulo_id))
		{			
			vModifArticulo_Abrir()
		}
		if(not aliasText){
			if(not WinExist(vProovedoresHabituales_id))
			{
				ControlClick, %vReporteArticulos_proovedoresHabituales%, %vModifArticulo_id%,,,, NA ;Clickea el boton Proveedores Habituales
				WinWait, %vProovedoresHabituales_id%, , 5
				if ErrorLevel {
					MsgBox, GetAlias - Could not rouse vProovedoresHabituales_id from the dead.
					return
				}
				ControlClick, TBTNBMP11, %vProovedoresHabituales_id%,,,, NA
				WinWait, %vVerProovedorHabitual_id%, , 5
				if ErrorLevel
				{
					MsgBox, GetAlias - Could not summon vModifProovedor_id to this mortal coil.
					return
				}
			}
			ControlGetText, aliasText, %vVerProovedorHabitual_alias%, %vVerProovedorHabitual_id%
			
			WinKill, %vVerProovedorHabitual_id%
			ControlClick, Salir, %vVerProovedorHabitual_id%,,,, NA
			WinKill, %vProovedoresHabituales_id%
		}
	}
	
	if(checkNota)
	{
		if(not WinExist(vModifArticulo_id))
		{
			vModifArticulo_Abrir()
		}
		ControlGetText, notaAdicional, %vModifArticulo_nota%, %vModifArticulo_id%
		RegExMatch(notaAdicional, "im).*(?:Alias completo|Alias|Simil):[ ]+(.*)$", aliasReplacement)
		if(aliasReplacement1)
		{
			aliasText := aliasReplacement1
		}		
	}
	
	if(parseAfter){
		aliasText := ParseAlias(aliasText)
	}
	return aliasText
}
;}

;{ Búsqueda
Buscar(){
	if(shouldStop())
	{
		return
	}
	working := true
			
	alias := GetAlias()
	alias := RegExReplace(alias, "\s+NO TRAER", "") ;Just remove these common words.
	if(alias == "NEXT!"){
		if(shouldStop())
		{
			return
		}
		ProximoArticulo(false)
		Buscar()
		return
	}
	if(not alias){
		return
	}
	
	if WinExist(vCalc_id)
	{
        WinActivate, %vCalc_id%
        WinWait, %vCalc_id%
		
        if WinExist(vCalc_buscar)
		{
            WinActivate, %vCalc_buscar%
            WinWait, %vCalc_buscar%
            Send, !s ;Alt+S: Search For
        }
        else
		{
            Send, ^f ;Ctrl+F: Buscar
        }
        WinWait, %vCalc_buscar%
		
        WinGet, ventBuscarID, ID, %vCalc_buscar%
		curCoords := Calc_GetSelectedCoords()
		
        SendRaw, % alias
		Send, {Del}
		Send, {Enter}
		Loop{
			Sleep, 50
			
			if(A_Cursor == "Wait")
			{
				Continue
			}
			
			if(curCoords != Calc_GetSelectedCoords())
			{
				OnSuccessfulSearch()
				return 1
			}
			
			if(not WinActive(ventBuscarID))
			{
				WinGet, activeID, ID, A
				GetClientSize(activeID, winWidth, winHeight)
				if(winHeight == 89){ ;"End of File" dialog
					Send, {Enter}
					Continue
				}
				if(winHeight == 85){ ;"Not Found" dialog
					Send, {Enter}
					OnUnsuccessfulSearch()
					return 0
				}			
			}
			
            if(A_Index = 20){
				OnUnsuccessfulSearch()
				return 0
            }
        }
		OnUnsuccessfulSearch()
		return 0
    }
	else if WinExist(vAdobeBuscar_id)
	{
		ControlGetText, notaAdicional, %vModifArticulo_nota%, %vModifArticulo_id%
		RegExMatch(notaAdicional, "im).*Pagina:[ ]+(.*)$", pageOverrides)
		if(pageOverrides1)
		{
			WinActivate, %vAdobe_id%
			WinWait, %vAdobe_id%
			Send, ^+n ;Ctrl+Shift+N: Go To Page
			Send, %pageOverrides1%{Enter}
			return
		}
		
        WinActivate, %vAdobeBuscar_id%
        WinWait, %vAdobeBuscar_id%
        ControlClick, %vAdobeBuscar_ok%, %vAdobeBuscar_id%
        WinWait, %vAdobeBuscar_id%
        ControlFocus, %vAdobeBuscar_input%, %vAdobeBuscar_id%
		ControlSetText, %vAdobeBuscar_input%,, %vAdobeBuscar_id%
		SendRaw, % alias
		Send, {Enter}
		
		WinWait, %vAdobeBuscar_id%
		WaitControlNotExist("Stop", vAdobeBuscar_id)
		WaitControlExist("Nueva búsqueda", vAdobeBuscar_id)
        WinWait, %vAdobeBuscar_id%
		
        ControlGetText, resultsText, %vAdobeBuscar_resultados%, %vAdobeBuscar_id%
        if(InStr(resultsText, "0 doc")){
            OnUnsuccessfulSearch()
            return 0
        }
		else if(InStr(resultsText, "1 instanc")){
			ControlClick, AVSearchTreeDocItemView, %vAdobeBuscar_id%
            OnSuccessfulSearch()
            return 1
        }
        else if(InStr(resultsText, "instanc")){
            OnSuccessfulSearch()
            return 1
        }
    }
	else if WinExist(vAdobe_id)
	{
        WinActivate, %vAdobe_id%
        WinWait, %vAdobe_id%
        Send, ^f ;Ctrl+F: Buscar
        Sleep, 100
		SendRaw, % alias
		Send, {Enter}
        return 1
    }
	else if WinExist(vWord_id)
	{
        WinActivate, %vWord_id%
        WinWait, %vWord_id%
        Send, ^b ;Ctrl+B: Buscar
        Sleep, 100
		SendRaw, % alias
		Send, {Enter}
        return 1
    }
	
	working := false	
}

OnUnsuccessfulSearch(){
	if(autoPilot) ;living on a edge baby
	{
		if(shouldStop){
			working := false
			shouldStop := false
		}
		else{
			working := true
			ProximoArticulo(false)
			Buscar()
		}
	}
}
OnSuccessfulSearch(){
	if(shouldStop())
	{
		return
	}
	
	ControlGetText, notaAdicional, %vModifArticulo_nota%, %vModifArticulo_id%
	RegExMatch(notaAdicional, "im).*Tooltip:[ ]+(.*)$", tooltips)
	if(tooltips1)
	{
		ToolTip, %tooltips1%
		SetTimer, RemoveToolTip, -1500
	}

	if WinExist(vCalc_id)
	{
		WinActivate, %vCalc_main%
		
		for index, match in AllRegexMatches(PostSearchString, "{[^{}]+}")
		{
			if(match == "{Seek}"){
				ControlGetText, oldPrice, %vModifArticulo_precioCosto%, %vModifArticulo_id%
				oldPrice := TextPrice2Float(oldPrice)
				ControlGetText, notaAdicional, %vModifArticulo_nota%, %vModifArticulo_id%
				
				RegExMatch(notaAdicional, "im).*Seek:[ ]+(.*)$", seekOverride)
				
				if(not seekOverride1 and forceSeek and Calc_IsInMergedCell())
				{
					ControlGetText, notaAdicional, %vModifArticulo_nota%, %vModifArticulo_id%
					InputBox, tempSeekInput, Nuevo Seek..., Ingrese el nuevo Seek.,,,,,,,,{End}{Left 2}
					if(ErrorLevel)
					{
						MsgBox, Seek inválido. (%tempSeekInput%)
						return
					}
					else
					{
						finalNota := "Seek: " . tempSeekInput . "`n" . notaAdicional
						WinActivate, %vModifArticulo_id%
						SetEdit(vModifArticulo_nota, vModifArticulo_id, finalNota)
						Sleep, 100
					}
					
				}
				
				;Seek: {End}{Down}{Left 2}
				if(seekOverride1)
				{
					if not(RegExMatch(seekOverride1, "{.*}")){
						seekOverride1 := "{" . seekOverride1 . "}"
					}
					Send, % seekOverride1
					Sleep, 100
					Send {Ctrl Down}c{Ctrl Up}
				}
				else
				{
					Clipboard := Calc_SeekInRow(oldPrice)
				}
			}
			else if(match == "{Paste}"){
				WinWait, A
				Send {Ctrl Down}c{Ctrl Up}
				WinWait, A
				success := PastePrice()
				if(success and (not shouldStop) and lastPercent <= 20 and lastPercent >= -15) ;living on a EEEEDGE
				{
					working := true
					Send, {Launch_Mail}
				}
				else
				{
					working := false
					shouldStop := false
				}
			}
			else{
				Send, % match
			}
		}
	}

}
;}

;{ Precios
TextPrice2Float(price){
	price := RegExReplace(price, "[^0-9.,]") ;Non-numbers begone. This includes you, whitespace. This includes you too, linebreaks.
	
	if(RegExMatch(price, "\d+\.\d{3}\,\d+")){ ;Example: 11.517,12
		price := RegExReplace(price, "\.") ;Remove separator dots
		price := RegExReplace(price, "\,", ".") ;Commas to something that actually makes sense
	}
	else if(RegExMatch(price, "\d+\,\d{3}\.\d+")){ ;Example: 11,517.12
		price := RegExReplace(price, "\,") ;Remove separator commas
	}
	else if(RegExMatch(price, "\d+\,\d+")){ ;Example: 11517,12
		price := RegExReplace(price, "\,", ".") ;Commas to something that actually makes sense
	}
	return price
}

ApplyPriceMultipliers(ByRef newPrice, byRef oldPrice := 0, ByRef modificadorAdicionalString := "", ByRef precioAdicionalString := ""){
	ControlGetText, notaAdicional, %vModifArticulo_nota%, %vModifArticulo_id%

	RegExMatch(notaAdicional, "im).*Incluye (.*)$", preciosAdicionales)
	if(preciosAdicionales1)
	{
		if(not preciosGuardados[preciosAdicionales1])
		{
			explanation := "Ingrese un precio para: " . preciosAdicionales1 . "`n(Sin ningún aumento o descuento, tal como aparece en la lista)"
			InputBox, tempInput, Guardar precio..., %explanation%
			if(ErrorLevel or not IsNum(tempInput))
			{
				MsgBox, Precio inválido. (%tempInput%)
				return
			}
			else
			{
				preciosGuardados[preciosAdicionales1] := tempInput
			}
		}
		precioAdicional := preciosGuardados[preciosAdicionales1]
		precioAdicionalString := " +" . precioAdicional
		newPrice := newPrice + precioAdicional
	}
	
	modificadorAdicionalString := ""
    RegExMatch(notaAdicional, "im).*Precio de lista \*([0-9.]+)$", extraMults)
    if(extraMults1)
	{
        modificadorAdicionalString := " *" . extraMults1
        newPrice := newPrice * extraMults1
    }
    RegExMatch(notaAdicional, "im).*Precio de lista \/([0-9.]+)$", extraDivisions)
    if(extraDivisions1)
	{
		modificadorAdicionalString := " /" . extraDivisions1
        newPrice := newPrice / extraDivisions1
    }
	
	newPrice := newPrice * modificadoresMult
	
	if(newPrice * 500 < oldPrice){ ;FUCK THOUSANDS SEPARATORS
        newPrice := newPrice * 1000
    }
}

PastePrice(newPrice := 0){
	if(shouldStop())
	{
		return
	}
	
	if(not WinExist(vModifArticulo_id))
	{
		vModifArticulo_Abrir()
	}
	
	ControlGetText, oldPrice, %vModifArticulo_precioCosto%, %vModifArticulo_id%
	oldPrice := TextPrice2Float(oldPrice)
	ControlGetText, itemID, %vModifArticulo_codigo%, %vModifArticulo_id%
	
	if(newPrice == 0)
	{
		newPrice := Clipboard
	}
	newPrice := TextPrice2Float(newPrice)
	if(not IsNum(newPrice))
	{
        MsgBox, Precio inválido. (%newPrice%)
        return
    }
	
	modificadorAdicionalString := ""
	precioAdicionalString := ""
	ApplyPriceMultipliers(newPrice, oldPrice, modificadorAdicionalString, precioAdicionalString)

	percent := (100*newPrice/oldPrice)-100
    percent := Round(percent, 1)
    if((percent < -15 or percent > 20) and not suppressWarnings)
	{
        MsgBox, 305, Diferencia de Precios, Diferencia de %percent%`%, continuar? ;1+48+256
        IfMsgBox, Cancel
        {
            return 0
        }
    }
	
	newPrice := Round(newPrice, 3) ;Lupa quiere 3 decimales.
	
	ControlFocus, %vModifArticulo_precioCosto%, %vModifArticulo_id%
	WinActivate, %vModifArticulo_id% ;TODO
	ControlSend, %vModifArticulo_precioCosto%, %newPrice%, %vModifArticulo_id%
	
	;Control, ChooseString, Dolares, %vModifArticulo_moneda%, %vModifArticulo_id%
	
	LogPriceChange(itemID, oldPrice, newPrice, modificadoresText, modificadorAdicionalString, precioAdicionalString)
	lastPercent := percent
	return true
}

SetMargins(margin1, margin2, margin3, fast := false){
	if(not WinExist(vModifArticulo_id))
	{
		vModifArticulo_Abrir()
	}
	
	ControlFocus, %vModifArticulo_margen1%, %vModifArticulo_id%
	Send, %margin1%{Enter}
	Sleep, 100
	ControlFocus, %vModifArticulo_margen2%, %vModifArticulo_id%
	Send, %margin2%{Enter}
	Sleep, 100
	ControlFocus, %vModifArticulo_margen3%, %vModifArticulo_id%
	Send, %margin3%{Enter}
	Sleep, 100
	ControlFocus, ComboBox4, %vModifArticulo_id%
	Control, ChooseString, VENTAS/COMPRAS, ComboBox4, %vModifArticulo_id%
	;ControlSend, ComboBox4, {Enter}, %vModifArticulo_id%
	
	if(fast)
	{
		Sleep, 100
		ControlFocus, Ok, %vModifArticulo_id%,,,, NA
		ControlClick, Ok, %vModifArticulo_id%,,,, NA
		ProximoArticulo(false)
	}
	
	
	return true
}

FastSetRubro(rubro){
	if(not WinExist(vModifArticulo_id))
	{
		vModifArticulo_Abrir()
	}
	
	ControlFocus, %vModifArticulo_rubro%, %vModifArticulo_id%,,,, NA
	Control, ChooseString, %rubro%, %vModifArticulo_rubro%, %vModifArticulo_id%
	Sleep, 100
	ControlFocus, Ok, %vModifArticulo_id%,,,, NA
	ControlClick, Ok, %vModifArticulo_id%,,,, NA
	Sleep, 200
	ProximoArticulo(false)
	
	return true
}
;}

;{ Navegación
ProximoArticulo(openAfter := true)
{
	if(shouldStop())
	{
		return
	}
		
	if(WinExist(vModifArticulo_id))
	{
		vModifArticulo_Cerrar()
	}
	
	ControlSend, %vReporteArticulos_planilla%, {Down}, %vReporteArticulos_id%
	if(openAfter)
	{
		vModifArticulo_Abrir()
	}
}

AnteriorArticulo(openAfter := true)
{
	if(shouldStop())
	{
		return
	}
		
	if(WinExist(vModifArticulo_id))
	{
		vModifArticulo_Cerrar()
	}
	ControlSend, %vReporteArticulos_planilla%, {Up}, %vReporteArticulos_id%
	if(openAfter)
	{
		vModifArticulo_Abrir()
	}
}
;}

;{ Logging
LogPriceChange(itemID := "", oldPrice := "", newPrice = "", modificadores := "", modificadorAdicional := "", precioAdicional := ""){
    percent := (100*newPrice/oldPrice)-100
    percent := Round(percent, 1)
    finalText = %itemID%: %percent%`% (%lastSeekCol%, %modificadores%%modificadorAdicional%%precioAdicional%, %oldPrice% -> %newPrice%)
    finalText = %finalText%`r`n ;concatenation
    LogSend(finalText)
}

LogSend(finalText := ""){
    if(not WinExist(vNotepad_id))
	{
        prev := WinActive("A")
        Run, Notepad
        WinWait, %vNotepad_id%
        WinActivate, ahk_id %prev%
    }
    ControlSend,,^{End}, %vNotepad_id% ;Ctrl+End: Go to end of document
    Control, EditPaste, %finalText%, , %vNotepad_id%
}
;}

;{ Opciones
Menu, Tray, Add  ; Add a separator line.
boundSetPostSearchString := Func("SetPostSearchString").Bind("", true)
Menu, Tray, Add, Post-Search Commands..., % boundSetPostSearchString
SetPostSearchString(searchStringInput := "", displayMessage := true)
{
	if(searchStringInput == "")
	{
		explanation := "Write out a set of instructions to send after a successful search.`r`rEach instruction must be between curly brackets, such as: {Right}`rAdd a number after your instruction to make it repeat that many times, for example: {Right 2}.`rSyntax is the same as AutoHotKey's Send command.`rSpecial commands: {Seek} and {Paste}."
		if(PostSearchString == "")
		{
			defaultInput := "{Seek}{Paste}"
		}
		else
		{
			defaultInput := PostSearchString
		}
		InputBox, searchStringInput, Post-Search Commands, %explanation%, , 420, 260, , , , , %defaultInput%
		if(ErrorLevel){
			return
		}
	}

    PostSearchString := searchStringInput
	RegWrite, REG_SZ, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedPostSearchString, %PostSearchString%
	if(displayMessage)
	{
		MsgBox, %PostSearchString%
	}
}


Menu, Tray, Add  ; Add a separator line.

; Create a submenu in the first menu (a right-arrow indicator). When the user selects it, the second menu is displayed.
Menu, Tray, Add, Search Type, :searchTypeMenu
Menu, searchTypeMenu, Add, Deactivate All, DeactivateAllSearchTypes

Menu, Tray, Add  ; Add a separator line.

Menu, Tray, Add, Suppress Warnings, toggleSuppressWarnings
toggleSuppressWarnings(){
    if(suppressWarnings == true){
        Menu, Tray, Uncheck, Suppress Warnings
        suppressWarnings := false
		RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSuppressWarnings
    }
    else{
        Menu, Tray, Check, Suppress Warnings
        suppressWarnings := true
		RegWrite, REG_SZ, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSuppressWarnings, Yes
    }
}

Menu, Tray, Add, Skip on Unsuccessful Search, toggleAutoPilot
toggleAutoPilot(){
    if(autoPilot == true){
        Menu, Tray, Uncheck, Skip on Unsuccessful Search
        autoPilot := false
		RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSkipOnSearch
    }
    else{
        Menu, Tray, Check, Skip on Unsuccessful Search
        autoPilot := true
		RegWrite, REG_SZ, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSkipOnSearch, Yes
    }
}

Menu, Tray, Add, Force Seek, toggleForceSeek
toggleForceSeek(){
    if(forceSeek == true){
        Menu, Tray, Uncheck, Force Seek
        forceSeek := false
		RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedForceSeek
    }
    else{
        Menu, Tray, Check, Force Seek
        forceSeek := true
		RegWrite, REG_SZ, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedForceSeek, Yes
    }
}

Menu, Tray, Add  ; Add a separator line.
Menu, Tray, Add, Exit, Exit
;}

;{ Misc
shouldStop(){
	if(shouldStop){
		working := false
		shouldStop := false
		return 1
	}
}

AllRegexMatches(haystack, needle){
    Pos := 1
    Matches := []
    M := ""
    while(Pos := RegExMatch(haystack, needle, M, Pos + StrLen(M)))
    {
        Matches.Push(M)
    }
    return Matches
}

GetClientSize(hWnd, ByRef w := "", ByRef h := "") ;Copied from Window Spy. Returns the "true" height/width of a window.
{
	VarSetCapacity(rect, 16)
	DllCall("GetClientRect", "ptr", hWnd, "ptr", &rect)
	w := NumGet(rect, 8, "int")
	h := NumGet(rect, 12, "int")
}

IsNum(str){ ;Fuck AHK.
	if str is number
		return true
	return false
}

GetClassName(hwnd)
{ ; returns HWND's class name without its instance number, e.g. "Edit" or "SysListView32"
	VarSetCapacity( buff, 256, 0 )
	DllCall("GetClassName", "uint", hwnd, "str", buff, "int", 255 )
	return buff
}

SetEdit(controlName, windowName, newText)
{
	ControlFocus, %controlName%, %windowName%
	Sleep, 50
	ControlSetText, %controlName%,, %windowName%
	Sleep, 50
	SendRaw, % newText
	Sleep, 100
}

DeepCopyControl(controlName, windowName1, windowName2, blacklist := "")
{
	ControlGet, controlHwnd, Hwnd,,%controlName%,%windowName1%
	controlType := GetClassName(controlHwnd)
	if(controlType == "Edit")
	{
		ControlGetText, controlText, %controlName%, %windowName1%
		if(blacklist){
			StringReplace, controlText, controlText, %blacklist%, , All
		}
		ControlFocus, %controlName%, %windowName2%
		ControlSetText, %controlName%,, %windowName2%
		SendRaw, % controlText
		Sleep, 200
		;Send, {Enter}
		
		;Control, EditPaste, %controlText%, %controlName%, %windowName2%
		;ControlSend, %controlName%, %controlText%, %windowName2%
		;ControlSend, %controlName%, %controlText%, %windowName2%
	}
	if(controlType == "ComboBox")
	{
		ControlGet, controlText, Choice, , %controlName%, %windowName1%
		ControlFocus, %controlName%, %windowName2%
		Control, ChooseString, %controlText%, %controlName%, %windowName2%
	}
}

WaitControlExist(controlName, windowName, tries := 100, retryTimer := 50){
	Loop
	{
		ControlGet, C, Visible,, %controlName%, %windowName%
		if(C)
		{
			return 1
		}
		else{
			Sleep, retryTimer
		}
	}
	Until A_Index > tries
}
WaitControlNotExist(controlName, windowName, tries := 100, retryTimer := 50){
	Loop
	{
		ControlGet, C, Visible,, %controlName%, %windowName%
		if(!C)
		{
			return 1
		}
		else{
			Sleep, retryTimer
		}
	}
	Until A_Index > tries
}

FirstWindowThatExists(windows){
	for index, windcandidate in windows
	{
		If(WinExist(windcandidate)){
			return windcandidate
		}
	}
}

WindowUnderMouse()
{
	MouseGetPos,,,underCursor
	WinGetTitle, Title, ahk_id %underCursor%
	Return, Title
}

Calc_GetSelectedCoords()
{
	oSM := ComObjCreate("com.sun.star.ServiceManager")			; This line is mandatory with AHK for OOo API
	oDesk := oSM.createInstance("com.sun.star.frame.Desktop")	; Create the first and most important service
	Array := ComObjArray(VT_VARIANT:=12, 2)
	Array[1] := MakePropertyValue(oSM, "Hidden", ComObject(0xB,true))
	oDoc := oDesk.CurrentComponent("private:factory/scalc", "_blank", 0, Array)  
	oSel := oDoc.getCurrentSelection
	oCell := ""
	if(oSel.getImplementationName == "ScCellObj"){
		oCell := oSel
	}
	else if(oSel.getImplementationName == "ScCellRangeObj"){
		oCell := oSel.getCellByPosition(0,0)
	}
	else if(oSel.getImplementationName == "ScCellRangesObj"){ ;SSSSSSSSSSSSSS
		oCell := oSel.getByIndex(0).getCellByPosition(0,0)
	}
	else{
		MsgBox % oSel.getImplementationName
		return 
	}
	Col:=oCell.CellAddress.Column
	Row:=oCell.CellAddress.Row 
	FinalStr := Col "-" Row
	Return FinalStr
}

Calc_IsInMergedCell()
{
	oSM := ComObjCreate("com.sun.star.ServiceManager")			; This line is mandatory with AHK for OOo API
	oDesk := oSM.createInstance("com.sun.star.frame.Desktop")	; Create the first and most important service
	Array := ComObjArray(VT_VARIANT:=12, 2)
	Array[1] := MakePropertyValue(oSM, "Hidden", ComObject(0xB,true))
	oDoc := oDesk.CurrentComponent("private:factory/scalc", "_blank", 0, Array)  
	oSel := oDoc.getCurrentSelection
	if(oSel.getImplementationName == "ScCellRangeObj"){
		return true
	}
	else{
		return false
	}
}

Calc_SeekInRow(theVal)
{
	oSM := ComObjCreate("com.sun.star.ServiceManager")			; This line is mandatory with AHK for OOo API
	oDesk := oSM.createInstance("com.sun.star.frame.Desktop")	; Create the first and most important service	
	Array := ComObjArray(VT_VARIANT:=12, 2)
	Array[1] := MakePropertyValue(oSM, "Hidden", ComObject(0xB,true))
	oDoc := oDesk.CurrentComponent("private:factory/scalc", "_blank", 0, Array)

	oDoc.getCurrentController.Select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) ;Deselect

	oSel := oDoc.getCurrentSelection
	oCell := ""
	oActiveSheet := oDoc.getCurrentController().getActiveSheet()
	
	if(oSel.getImplementationName == "ScCellObj"){
		oCell := oSel
	}
	else if(oSel.getImplementationName == "ScCellRangeObj"){
		oCell := oSel.getCellByPosition(0,0)
	}
	else if(oSel.getImplementationName == "ScCellRangesObj"){ ;SSSSSSSSSSSSSS
		oCell := oSel.getByIndex(0).getCellByPosition(0,0)
	}
	else{
		MsgBox % oSel.getImplementationName
		return 
	}
	
	closestNumbr := 0
	searchResultRow := oCell.CellAddress.Row
	searchResultCol := oCell.CellAddress.Column
	candidateRow := searchResultRow
	candidateCol := searchResultCol
	
	;Get the used range of the sheet, first, so we can use it as reference for what columns to evaluate.
	oCursor := oActiveSheet.createCursor()
	oCursor.gotoStartOfUsedArea(False)
	oCursor.gotoEndOfUsedArea(True)
	oUsedRange := oCursor.getRangeAddress()
	;Get the area of our base cell, whether merged or not, so we can use it as for reference for what rows to evaluate.
	oCursor := oActiveSheet.createCursorByRange(oCell)
	oCursor.collapseToMergedArea()
	oMergedArea := oCursor.getRangeAddress()
	;All together now!
	rg := oActiveSheet.getCellRangeByPosition(oUsedRange.StartColumn, oMergedArea.StartRow, oUsedRange.EndColumn, oMergedArea.EndRow)
	mData := rg.getDataArray()
	
	row_ := -1
	while(row_ < mData.MaxIndex()){
		row_++
		col_ := -1
		for key in mData[row_]
		{
			col_++
			currRow := searchResultRow + row_
			currCol := oUsedRange.StartColumn + col_
			;MsgBox, Evaluating %currRow% - %currCol%, which has a value of %key%!
			key := RegExReplace(key, "\s+USD", "") ;Just remove these common words.
			if(RegExMatch(key, "(?:[a-zA-Z]+[0-9.,]|[0-9.,]+[a-zA-Z])[a-zA-Z0-9.,]*")) ;I genuinely have no idea what this is.
			{
				continue
			}
			if(RegExMatch(key, "[^0-9 ]{5,}")) ;No more than 5 non-number characters please
			{
				continue
			}
			if(RegExMatch(key, "\d[^\d\n\.\,\d]+\d")) ;Numbers shall not be intersected by anything besides a dot or a comma (1x20 is bad, 17 18 is bad etc)
			{
				continue
			}
			numbr := TextPrice2Float(key)
			if(IsNum(numbr))
			{
				ApplyPriceMultipliers(numbr, theVal)
				;todo: apply last good modifier
				if(abs(theVal-numbr) <= abs(theVal-closestNumbr))
				{
					;MsgBox, %currRow% - %currCol%: %theVal%-%numbr% / %theVal%-%closestNumbr%
					closestNumbr := numbr
					candidateRow := currRow
					candidateCol := currCol
					lastSeekCol := col_ - searchResultCol
				}
			}
		}
	}
	;MsgBox, Selecting %candidateCol% - %candidateRow%!
	oCell := oActiveSheet.getCellByPosition(candidateCol,candidateRow)
	oDoc.getCurrentController().Select(oCell)
	oDoc.getCurrentController.Select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) ;Deselect
	
	return closestNumbr
}

;Used for OOo API.
MakePropertyValue(poSM, cName, uValue)
{	oPropertyValue			:= Object()
	oPropertyValue 			:= poSM.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
	oPropertyValue.Name		:= cName
	oPropertyValue.Value	:= uValue
	Return oPropertyValue
}
;}

;{ AUTOEXEC
RegRead, savedModificadores, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedModificadores
RegRead, savedPostSearchString, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedPostSearchString
RegRead, savedSearchTypes, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSearchTypes
RegRead, savedSuppressWarnings, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSuppressWarnings
RegRead, savedSkipOnSearch, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSkipOnSearch
RegRead, savedForceSeek, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedForceSeek
if(savedModificadores or savedPostSearchString or savedSearchTypes or savedSuppressWarnings or savedSkipOnSearch or savedForceSeek)
{
	explanationConfig := "Saved config:"
	if(savedModificadores)
	{
		explanationConfig = %explanationConfig%`nModificadores: %savedModificadores%
	}
	if(savedPostSearchString)
	{
		explanationConfig = %explanationConfig%`nPost-Search String: %savedPostSearchString%
	}
	if(savedSearchTypes)
	{
		stringWithNoLinebreaks := RTrim(StrReplace(savedSearchTypes, "`n", ", "), ", ")
		explanationConfig = %explanationConfig%`nActive Search Types: %stringWithNoLinebreaks%
	}
	if(savedSuppressWarnings)
	{
		explanationConfig = %explanationConfig%`nSuppress Warnings: %savedSuppressWarnings%
	}
	if(savedSkipOnSearch)
	{
		explanationConfig = %explanationConfig%`nSkip on Unsuccessful Search: %savedSkipOnSearch%
	}
	if(savedForceSeek)
	{
		explanationConfig = %explanationConfig%`nForceSeek: %savedForceSeek%
	}
	explanationConfig = %explanationConfig%`n`nImport?

	MsgBox, 305, Import Config, %explanationConfig% ;1+48+256
	IfMsgBox, OK
	{
		if(savedModificadores)
		{
			SetModificadores(savedModificadores, false)
		}
		if(savedPostSearchString)
		{
			SetPostSearchString(savedPostSearchString, false)
		}
		if(savedSearchTypes)
		{
			for i, currSearchType in searchTypes {
				if InStr(savedSearchTypes, currSearchType.name . "`n")
					currSearchType.active := true
				else
					currSearchType.active := false
			}
			refreshSearchTypeMenu()
		}
		if(savedSuppressWarnings)
		{
			toggleSuppressWarnings()
		}
		if(savedSkipOnSearch)
		{
			toggleAutoPilot()
		}
		if(savedForceSeek)
		{
			toggleForceSeek()
		}
	}
	IfMsgBox, Cancel
	{
		RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedModificadores
		RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedPostSearchString
		RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSearchTypes
		RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSuppressWarnings
		RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSkipOnSearch
		RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedForceSeek
	}
}
;}

;{ Keybinds
Launch_Media::
;FastCorrectNota("Monteluz")
;FastSetRubro("16")
;SetMargins("60", "25", "40", true)
;Msgbox, Testing...	
;FastAliasizeDesc()
WinRestore, LUPA - Gest

return

!^Launch_Media::
ListLines 
return

Volume_Up::
return

Volume_Down::
return

Volume_Mute::
SetModificadores()
return

Media_Play_Pause::
return

Media_Prev::
;MouseClick, X1
;return
AnteriorArticulo(false)
Buscar()
return

Media_Next::
;MouseClick, X2
;return
ProximoArticulo(false)
Buscar()
return

Launch_Mail::
If(WinExist("Diferencia"))
{
	Send, {Left}{Enter}
	return
}
if WinExist(vFacturaProov_id)
{ ;todo move this to functions up top
	if (WinExist(vFacturaProovModif_id))
	{
		ControlGet, tipoIVA, Choice, , ComboBox1, %vFacturaProovModif_id%
		if (InStr(tipoIVA, "NETO GRAVADO"))
		{
			WinMove, 100, 100
			ControlGetText, netoPrevio, Edit1, %vFacturaProovModif_id%
			ControlGetText, netoReal, Static12, %vFacturaProov_id%
			InputBox, totalTemp, Importe Sumado, Ingrese el importe sumado que aparece abajo a la derecha`, al lado del botón Asigna.
			if(ErrorLevel or not IsNum(totalTemp))
			{
				return
			}
			newNeto := TextPrice2Float(netoPrevio) - TextPrice2Float(totalTemp) + TextPrice2Float(netoReal)
			ControlFocus, Edit1, %vFacturaProovModif_id%
			ControlSetText, Edit1,, %vFacturaProovModif_id%
			SendRaw, % newNeto
			return
		}
	}

	primeraVentFact := FirstWindowThatExists([vFacturaProovModif_id, vFacturaProovNuevo_id])
	if (primeraVentFact)
	{
		ControlGetText, factCodigo, Edit1, %primeraVentFact%
		factCodigo := Trim(factCodigo)
		ControlGetText, factCantidad, Edit3, %primeraVentFact%
		StringReplace, factCantidad, factCantidad, ",", , All
		StringReplace, factCantidad, factCantidad, " ", , All
		factCantidad := RegExReplace(factCantidad,"(\.\d*?)0*$","$1")
		factCantidad := RegExReplace(factCantidad,"\.$")
		;RegExMatch(factCantidad, "([A-Z]+-\d+)", factCantidad)
		;factCantidad := %factFantidad%%A_Tab%
		ControlSend, Edit3, {Enter}, %primeraVentFact%
		ControlGetText, factPrecioCosto, Edit4, %primeraVentFact%
		ControlSend, Edit4, {Enter}, %primeraVentFact%
		
		WinWait, %vModifArticulo_id%
		
		ControlGetText, factNombreCompleto, %vModifArticulo_descripcion%, %vModifArticulo_id%
		factNombreCompleto := Trim(factNombreCompleto)
		factPrecioCosto := TextPrice2Float(factPrecioCosto)
		ApplyPriceMultipliers(factPrecioCosto)
		factPrecioCosto := RegExReplace(factPrecioCosto,"(\.\d*?)0*$","$1")
		factPrecioCosto := RegExReplace(factPrecioCosto,"\.$")
		
		ControlClick, %vReporteArticulos_proovedoresHabituales%, %vModifArticulo_id%,,,, NA ;Clickea el boton Proveedores Habituales
		WinWait, %vProovedoresHabituales_id%, , 5
		if ErrorLevel {
			MsgBox, GetAlias - Could not rouse vProovedoresHabituales_id from the dead.
			return
		}
		ControlClick, TBTNBMP11, %vProovedoresHabituales_id%,,,, NA
		WinWait, %vVerProovedorHabitual_id%, , 5
		if ErrorLevel
		{
			MsgBox, GetAlias - Could not summon vVerProovedorHabitual_id to this mortal coil.
			return
		}
		ControlGetText, factAliasText, %vVerProovedorHabitual_alias%, %vVerProovedorHabitual_id%
		factAliasText := Trim(factAliasText)
		ControlClick, Salir, %vVerProovedorHabitual_id%,,,, NA
		ControlSend,, {Esc}, %vProovedoresHabituales_id%
		vModifArticulo_Cerrar()
		ControlFocus, TWBROWSE1, %vFacturaProov_id%
		ControlSend, TWBROWSE1, {PGDN}, %vFacturaProov_id%
		
		finalDetailText = %factCantidad% x %factCodigo% (%factAliasText%) - %factPrecioCosto% - %factNombreCompleto%`r`n
		LogSend(finalDetailText)
		Sleep, 200
		ControlClick, Button5, %vFacturaProov_id%,,,, NA

	}
	return
}

WinWait, %vModifArticulo_id%
ControlSend, Ok, {Space}, %vModifArticulo_id%
ControlSend, Sí, {Enter}, Atención
ControlSend, Sí, {Enter}, Atención
ProximoArticulo(false)
Buscar()
return

!^Launch_Mail::
WinActivate, %vModifArticulo_id%
WinActivate, %vNuevoArticulo_id%
camposAClonar := [vModifArticulo_descripcion, vModifArticulo_puntoPedido, vModifArticulo_empaque, vModifArticulo_unidad, vModifArticulo_moneda, vModifArticulo_margen1, vModifArticulo_margen2, vModifArticulo_margen3, vModifArticulo_iva, vModifArticulo_rubro, vModifArticulo_nota]

;DeepCopyControl(vModifArticulo_precioCosto, vModifArticulo_id, vNuevoArticulo_id, ",")
for i, elCampo in camposAClonar
{
	DeepCopyControl(elCampo, vModifArticulo_id, vNuevoArticulo_id)
}
return

Browser_Search::
Buscar()
return

Browser_Home::
PastePrice()
return

^Browser_Home::
return

#If overrideMiddleClick
MButton::
If(InStr(WindowUnderMouse(), vReporteArticulos_id))
{
	vModifArticulo_Cerrar()
}
Else If(InStr(WindowUnderMouse(), vAdobe_id) or InStr(WindowUnderMouse(), vCalc_id) or InStr(WindowUnderMouse(), vWord_id))
{
	Click, 2
	Sleep, 500
	WinWait, A
	Send {Ctrl Down}c{Ctrl Up}
	Sleep, 300
	WinWait, A
	Send {Browser_Home}
}
Else
{
	Send {MButton}
}
return
#If

;#If working
;Esc::
;	shouldStop := true
;return
;#If
Pause::
	shouldStop := true
return
Scrolllock::
	shouldStop := true
return

RemoveToolTip:
ToolTip
return

Exit:
ExitApp
return