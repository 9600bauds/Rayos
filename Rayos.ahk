#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance force ; Close old versions of this script automatically.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2 ; Match window titles anywhere, not just at the start.

;{ Globals (most of these are effectively defines)
global ventModificarArticulo := "Modificación : "
global campoCodigo := "Edit1"
global campoPrecioCosto := "Edit7"
global campoNota := "Edit17"
global campoMargen1 := "Edit8"
global campoMargen2 := "Edit11"
global campoMargen3:= "Edit14"
global campoRubro := "ComboBox3"

global ventModificarProveedor := "Modificación"
global ventVerProveedor := "Consulta"
global ventProveedoresHabituales := "Proveedores Habituales"
global botonProovedoresHabituales := "Button7"
global campoAlias_Habituales := "Edit2"
global ventAliasProveedor := "Alias del Proveedor"
global campoAlias_Dedicado := "Edit1"

global ventReporteArticulos := "Artículos a Modificar"
global ventListaArticulos := "ARTICULOS-LA CASA DEL ELECTRICISTA"
global campoListado := "TXBROWSE1"

global ventNotepad := "ahk_class Notepad"
global ventWord := " - Word"
global ventCalc := "OpenOffice Calc"
global ventCalc_Buscar := "Find & Replace"
global ventCalc_Main := "ahk_class SALFRAME" ;Precisamente la planilla principal, no ningún diálogo
global ventAdobeReader := "Adobe Acrobat Reader"
global ventAdobeReader_Buscar := "ahk_class AVL_AVWindow"
global ventAdobeReader_BuscarOK := "Button18"
global ventAdobeReader_Buscar_Input := "Edit5"
global ventAbodeReader_Buscar_Matches := "Static12"

global ventFactProov := "FACTURA  Proveedor.Nueva"

global search_Default = "Default"
global search_Exact = "Exact"
global search_Start = "Match Start"
global search_End = "Match End"
global search_WordBoundaries = "Match Word Boundaries"
global search_RemoveLastWord = "Remove Last Word"
global search_RemoveLetters = "Remove Letters"
global search_LongestNumber = "Longest Number"
global search_LongestWord = "Longest Word"
global search_RemoveZeroes = "Remove ALL Trailing Zeroes"
global search_Fabrimport = "Fabrimport"
global search_Faroluz = "Faroluz"
global search_Ferrolux = "Ferrolux"
global search_Solnic = "Solnic"
global searchType := "Default"

global suppressWarnings := false
global autoPilot := false
global overrideMiddleClick := true

global modificadoresText := "+0%" ;These two should be equivalent and are only set once in SetModificadores().
global modificadoresMult := 1

global lastPercent := 0
global preciosGuardados := {}

global PostSearchString := ""

global working := false
global shouldStop := false
;}

;{ Ventana Lista Artículos

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
	
	if(InStr(alias, "NO TRAER") or InStr(alias, "NO COMPRAR") or RegExMatch(alias, "^[-]+$")){
		return "NEXT!"
	}
	
	if(searchType == search_Exact){
		alias := "^ " . alias . "$"
	}
	else if(searchType == search_Start){
		alias := "^ " . alias
	}
	else if(searchType == search_End){
		alias := alias . "$"
	}
	else if(searchType == search_WordBoundaries){
		alias := "\b" . alias . "\b"
	}
	else if(searchType == search_RemoveLastWord){
		alias := RegExReplace(alias, " \w+$", "")
	}
	else if(searchType == search_RemoveLetters){
		alias := RegExReplace(alias, "[A-Za-z]", "")
	}
	else if(searchType == search_LongestNumber){
		longestMatch := ""
		for index, match in AllRegexMatches(alias, "[\d]+"){
			if(StrLen(match) > StrLen(longestMatch)){
				longestMatch := match
			}
		}
		alias := longestMatch
	}
	else if(searchType == search_longestWord){
		longestMatch := ""
		for index, match in AllRegexMatches(alias, "[\w]+"){
			if(StrLen(match) > StrLen(longestMatch)){
				longestMatch := match
			}
		}
		alias := "\b" . longestMatch . "\b"
	}
	else if(searchType == search_removeZeroes){
		alias := RegExReplace(alias, "^[0]+", "")
	}
	else if(searchType == search_Fabrimport){
		alias := "[^0-9]" . alias . "$"
	}
	else if(searchType == search_Faroluz){
		alias := RegExReplace(alias, " \w+$", "") . "$"
	}
	else if(searchType == search_Ferrolux){
		RegExMatch(alias, "([A-Z]+-\d+)", alias)
		if WinExist(ventCalc)
		{
			alias := "^ " . alias
		} 
	}
	else if(searchType == search_Solnic){
		alias := "^" . alias . "[\s+|$]"
	}
	
	return alias
}

GetAlias(parseAfter := true, checkNota := true){
	if(shouldStop())
	{
		return
	}
		
	aliasText := ""
	if(WinExist(ventReporteArticulos))
	{
		if(WinExist(ventAliasProveedor)){
			ControlGetText, aliasText, %campoAlias_Dedicado%, %ventAliasProveedor%
			WinKill, %ventAliasProveedor%
		}
		if(not WinExist(ventModificarArticulo))
		{
			SetControlDelay -1
			Loop{
				if(A_Index = 20){
					MsgBox, GetAlias - Could not cometh here ventModificarArticulo.
					return
				}
				ControlClick, Modifica, %ventReporteArticulos%,,,, NA
				ControlSend, Modifica, {Space}, %ventReporteArticulos%
				WinWait, %ventModificarArticulo%, , 0.5
				if not ErrorLevel {
					Break
				}
			}
		}
		if(not aliasText){
			if(not WinExist(ventProveedoresHabituales))
			{
				ControlClick, %botonProovedoresHabituales%, %ventModificarArticulo%,,,, NA ;Clickea el boton Proveedores Habituales
				WinWait, %ventProveedoresHabituales%, , 5
				if ErrorLevel {
					MsgBox, GetAlias - Could not rouse ventProveedoresHabituales from the dead.
					return
				}
				ControlClick, Modifica, %ventProveedoresHabituales%,,,, NA
				WinWait, %ventModificarProveedor%, , 5
				if ErrorLevel
				{
					MsgBox, GetAlias - Could not summon ventModificarProveedor to this mortal coil.
					return
				}
			}
			ControlGetText, aliasText, %campoAlias_Habituales%, %ventModificarProveedor%
			
			WinKill, %ventModificarProveedor%
			WinKill, %ventProveedoresHabituales%
		}
	}
	
	if(checkNota)
	{
		if(not WinExist(ventModificarArticulo))
		{
			ControlClick, Modifica, %ventReporteArticulos%,,,, NA
			ControlSend, Modifica, {Space}, %ventReporteArticulos%
			WinWait, %ventModificarArticulo%, , 5
			if ErrorLevel {
				MsgBox, GetAlias - Could not cometh here ventModificarArticulo.
				return
			}
		}
		ControlGetText, notaAdicional, %campoNota%, %ventModificarArticulo%
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

FastSetAlias(thealias := ""){
	if(!thealias or theAlias = "")
		Send {Ctrl Down}c{Ctrl Up}
		Sleep, 100
		thealias := Clipboard
	ControlClick, TBTNBMP29, %ventReporteArticulos%
	WinWait, %ventAliasProveedor%
	WinActivate, %ventAliasProveedor%
	ControlFocus, Edit1, %ventAliasProveedor%
	Send, {Shift Down}{End}{Shift Up}
	thealias := "0" . thealias
	Sleep, 100
	SendRaw, % thealias
	Sleep, 100
	Send, {Enter 3}
	Sleep, 500
	ProximoArticulo(false)
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
	if(alias == "NEXT!"){
		ProximoArticulo(false)
		Buscar()
		return
	}
	if(not alias){
		return
	}
	
	if WinExist(ventCalc)
	{
        WinActivate, %ventCalc%
        WinWait, %ventCalc%
		
        if WinExist(ventCalc_Buscar)
		{
            WinActivate, %ventCalc_Buscar%
            WinWait, %ventCalc_Buscar%
            Send, !s ;Alt+S: Search For
        }
        else
		{
            Send, ^f ;Ctrl+F: Buscar
        }
        WinWait, %ventCalc_Buscar%
		
        WinGet, ventBuscarID, ID, %ventCalc_Buscar%
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
				if(winHeight == 74){ ;"End of File" dialog
					Send, {Enter}
					Continue
				}
				if(winHeight == 80){ ;"Not Found" dialog
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
	else if WinExist(ventAdobeReader_Buscar)
	{
		ControlGetText, notaAdicional, %campoNota%, %ventModificarArticulo%
		RegExMatch(notaAdicional, "im).*Pagina:[ ]+(.*)$", pageOverrides)
		if(pageOverrides1)
		{
			WinActivate, %ventAdobeReader%
			WinWait, %ventAdobeReader%
			Send, ^+n ;Ctrl+Shift+N: Go To Page
			Send, %pageOverrides1%{Enter}
			return
		}
		
        WinActivate, %ventAdobeReader_Buscar%
        WinWait, %ventAdobeReader_Buscar%
        ControlClick, %ventAdobeReader_BuscarOK%, %ventAdobeReader_Buscar%
        WinWait, %ventAdobeReader_Buscar%
        ControlFocus, %ventAdobeReader_Buscar_Input%, %ventAdobeReader_Buscar%
		ControlSetText, %ventAdobeReader_Buscar_Input%,, %ventAdobeReader_Buscar%
		SendRaw, % alias
		Send, {Enter}
		
		WinWait, %ventAdobeReader_Buscar%
		WaitControlNotExist("Stop", ventAdobeReader_Buscar)
		WaitControlExist("Nueva búsqueda", ventAdobeReader_Buscar)
        WinWait, %ventAdobeReader_Buscar%
		
        ControlGetText, resultsText, %ventAbodeReader_Buscar_Matches%, %ventAdobeReader_Buscar%
        if(InStr(resultsText, "0 doc")){
            OnUnsuccessfulSearch()
            return 0
        }
		else if(InStr(resultsText, "1 instanc")){
			ControlClick, AVSearchTreeDocItemView, %ventAdobeReader_Buscar%
            OnSuccessfulSearch()
            return 1
        }
        else if(InStr(resultsText, "instanc")){
            OnSuccessfulSearch()
            return 1
        }
    }
	else if WinExist(ventAdobeReader)
	{
        WinActivate, %ventAdobeReader%
        WinWait, %ventAdobeReader%
        Send, ^f ;Ctrl+F: Buscar
        Sleep, 100
		SendRaw, % alias
		Send, {Enter}
        return 1
    }
	else if WinExist(ventWord)
	{
        WinActivate, %ventWord%
        WinWait, %ventWord%
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
	
	ControlGetText, notaAdicional, %campoNota%, %ventModificarArticulo%
	RegExMatch(notaAdicional, "im).*Tooltip:[ ]+(.*)$", tooltips)
	if(tooltips1)
	{
		ToolTip, %tooltips1%
		SetTimer, RemoveToolTip, -1500
	}

	if WinExist(ventCalc)
	{
		WinActivate, %ventCalc_Main%
		
		for index, match in AllRegexMatches(PostSearchString, "{[^{}]+}")
		{
			if(match == "{Seek}"){
				ControlGetText, oldPrice, %campoPrecioCosto%, %ventModificarArticulo%
				oldPrice := TextPrice2Float(oldPrice)
				ControlGetText, notaAdicional, %campoNota%, %ventModificarArticulo%
				
				RegExMatch(notaAdicional, "im).*Seek:[ ]+(.*)$", seekOverride)
				if(seekOverride1)
				{
					seekOverride := "{" . seekOverride1 . "}"
					Send, % seekOverride
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
	ControlGetText, notaAdicional, %campoNota%, %ventModificarArticulo%

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
		if(newPrice * 10 < precioAdicional){ ;FUCK THOUSANDS SEPARATORS
			precioAdicional := precioAdicional / 1000
		}
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
	
	if(not WinExist(ventModificarArticulo))
	{
		ControlSend, Modifica, {Space}, %ventReporteArticulos%
		WinWait, %ventModificarArticulo%, , 5
		if ErrorLevel {
			MsgBox, PastePrice - Could not bring forth ventModificarArticulo.
			return
		}
	}
	
	ControlGetText, oldPrice, %campoPrecioCosto%, %ventModificarArticulo%
	oldPrice := TextPrice2Float(oldPrice)
	ControlGetText, itemID, %campoCodigo%, %ventModificarArticulo%
	
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
	
	ControlFocus, %campoPrecioCosto%, %ventModificarArticulo%
	WinActivate, %ventModificarArticulo% ;TODO
	ControlSend, %campoPrecioCosto%, %newPrice%, %ventModificarArticulo%
	
	LogPriceChange(itemID, oldPrice, newPrice, modificadoresText, modificadorAdicionalString, precioAdicionalString)
	lastPercent := percent
	return true
}

SetMargins(margin1, margin2, margin3, fast := false){
	if(not WinExist(ventModificarArticulo))
	{
		ControlSend, Modifica, {Space}, %ventReporteArticulos%
		WinWait, %ventModificarArticulo%, , 5
		if ErrorLevel {
			MsgBox, PastePrice - Could not special summon ventModificarArticulo.
			return
		}
	}
	
	ControlFocus, %campoMargen1%, %ventModificarArticulo%
	Send, %margin1%{Enter}
	Sleep, 100
	ControlFocus, %campoMargen2%, %ventModificarArticulo%
	Send, %margin2%{Enter}
	Sleep, 100
	ControlFocus, %campoMargen3%, %ventModificarArticulo%
	Send, %margin3%{Enter}
	Sleep, 100
	ControlFocus, ComboBox4, %ventModificarArticulo%
	Control, ChooseString, VENTAS/COMPRAS, ComboBox4, %ventModificarArticulo%
	;ControlSend, ComboBox4, {Enter}, %ventModificarArticulo%
	
	if(fast)
	{
		Sleep, 100
		ControlFocus, Ok, %ventModificarArticulo%,,,, NA
		ControlClick, Ok, %ventModificarArticulo%,,,, NA
		ProximoArticulo(false)
	}
	
	
	return true
}

FastSetRubro(rubro){
	if(not WinExist(ventModificarArticulo))
	{
		ControlSend, Modifica, {Space}, %ventReporteArticulos%
		WinWait, %ventModificarArticulo%, , 5
		if ErrorLevel {
			MsgBox, PastePrice - Could not blueprint construct ventModificarArticulo.
			return
		}
	}
	
	ControlFocus, %campoRubro%, %ventModificarArticulo%,,,, NA
	Control, ChooseString, %rubro%, %campoRubro%, %ventModificarArticulo%
	Sleep, 100
	ControlFocus, Ok, %ventModificarArticulo%,,,, NA
	ControlClick, Ok, %ventModificarArticulo%,,,, NA
	Sleep, 200
	ProximoArticulo(false)
	
	
	return true
}

FastCorrectNota(proov){ ;Unashamedly adhoc.
	if(proov == "Monteluz")
	{
		alias := GetAlias(false)
		ControlGetText, notaAdicional, %campoNota%, %ventModificarArticulo%
		if(RegExMatch(alias, "\-1") and !RegExMatch(notaAdicional, "im).*Seek:[ ]+Right 4$"))
		{
			SetEdit(campoNota, ventModificarArticulo, "Seek: Right 4")
		}
		else if(RegExMatch(alias, "\-7") and !RegExMatch(notaAdicional, "im).*Seek:[ ]+Right 6$"))
		{
			SetEdit(campoNota, ventModificarArticulo, "Seek: Right 6")
		}
		else if(RegExMatch(alias, "\-3") and !RegExMatch(notaAdicional, "im).*Seek:[ ]+Right 10$"))
		{
			SetEdit(campoNota, ventModificarArticulo, "Seek: Right 10")
		}
		else
		{
			MsgBox, Idk lol
			return
		}
		ControlFocus, Ok, %ventModificarArticulo%,,,, NA
		ControlClick, Ok, %ventModificarArticulo%,,,, NA
		Sleep, 200
		ProximoArticulo(false)
		;MsgBox, % alias
	}
}
;}

;{ Navegación
ProximoArticulo(openAfter := true)
{
	if(shouldStop())
	{
		return
	}
		
	if(WinExist(ventModificarArticulo))
	{
		WinKill, %ventModificarArticulo%
	}
	ControlSend, Sí, {Enter}, Atención
	ControlSend, Sí, {Enter}, Atención
	ControlSend, %campoListado%, {Down}, %ventReporteArticulos%
	if(openAfter)
	{
		ControlClick, Modifica, %ventReporteArticulos%,,,, NA
	}
}

AnteriorArticulo(openAfter := true)
{
	if(shouldStop())
	{
		return
	}
		
	if(WinExist(ventModificarArticulo))
	{
		WinKill, %ventModificarArticulo%
	}
	ControlSend, %campoListado%, {Up}, %ventReporteArticulos%
	if(openAfter)
	{
		ControlClick, Modifica, %ventReporteArticulos%,,,, NA
	}
}
;}
 
;{ Logging
LogPriceChange(itemID := "", oldPrice := "", newPrice = "", modificadores := "", modificadorAdicional := "", precioAdicional := ""){
    percent := (100*newPrice/oldPrice)-100
    percent := Round(percent, 1)
    finalText = %itemID%: %percent%`% (%modificadores%%modificadorAdicional%%precioAdicional%, %oldPrice% -> %newPrice%)
    finalText = %finalText%`r`n ;concatenation
    LogSend(finalText)
}

LogSend(finalText := ""){
    if(not WinExist(ventNotepad))
	{
        prev := WinActive("A")
        Run, Notepad
        WinWait, %ventNotepad%
        WinActivate, ahk_id %prev%
    }
    ControlSend,,^{End}, %ventNotepad% ;Ctrl+End: Go to end of document
    Control, EditPaste, %finalText%, , %ventNotepad%
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
;{
Menu, searchTypeMenu, Add, %search_Default%, setSearchDefault, Radio
setSearchDefault(){
    setSearchType(search_Default)
}
Menu, searchTypeMenu, Add, %search_Exact%, setSearchExact, Radio
setSearchExact(){
    setSearchType(search_Exact)
}
Menu, searchTypeMenu, Add, %search_Start%, setSearchStart, Radio
setSearchStart(){
    setSearchType(search_Start)
}
Menu, searchTypeMenu, Add, %search_End%, setSearchEnd, Radio
setSearchEnd(){
    setSearchType(search_End)
}
Menu, searchTypeMenu, Add, %search_wordBoundaries%, setSearchWordBoundaries, Radio
setSearchWordBoundaries(){
    setSearchType(search_wordBoundaries)
}
Menu, searchTypeMenu, Add, %search_RemoveLastWord%, setSearchRemoveLastWord, Radio
setSearchRemoveLastWord(){
    setSearchType(search_RemoveLastWord)
}
Menu, searchTypeMenu, Add, %search_RemoveLetters%, setSearchRemoveLetters, Radio
setSearchRemoveLetters(){
    setSearchType(search_RemoveLetters)
}
Menu, searchTypeMenu, Add, %search_LongestNumber%, setSearchLongestNumber, Radio
setSearchLongestNumber(){
    setSearchType(search_LongestNumber)
}
Menu, searchTypeMenu, Add, %search_LongestWord%, setSearchLongestWord, Radio
setSearchLongestWord(){
    setSearchType(search_LongestWord)
}
Menu, searchTypeMenu, Add, %search_removeZeroes%, setSearchRemoveZeroes, Radio
setSearchRemoveZeroes(){
    setSearchType(search_removeZeroes)
}
Menu, searchTypeMenu, Add, %search_Fabrimport%, setSearchFabrimport, Radio
setSearchFabrimport(){
    setSearchType(search_Fabrimport)
}
Menu, searchTypeMenu, Add, %search_Faroluz%, setSearchFaroluz, Radio
setSearchFaroluz(){
    setSearchType(search_Faroluz)
}
Menu, searchTypeMenu, Add, %search_Ferrolux%, setSearchFerrolux, Radio
setSearchFerrolux(){
    setSearchType(search_Ferrolux)
}
Menu, searchTypeMenu, Add, %search_Solnic%, setSearchSolnic, Radio
setSearchSolnic(){
    setSearchType(search_Solnic)
}
;}
setSearchType(search_Default, true)

setSearchType(type, initial := false){
    searchType := type
    
    Menu, searchTypeMenu, Uncheck, %search_Default%
    Menu, searchTypeMenu, Uncheck, %search_Exact%
    Menu, searchTypeMenu, Uncheck, %search_Start%
    Menu, searchTypeMenu, Uncheck, %search_End%
	Menu, searchTypeMenu, Uncheck, %search_WordBoundaries%
    Menu, searchTypeMenu, Uncheck, %search_RemoveLastWord%
	Menu, searchTypeMenu, Uncheck, %search_RemoveLetters%
    Menu, searchTypeMenu, Uncheck, %search_LongestNumber%
	Menu, searchTypeMenu, Uncheck, %search_LongestWord%
	Menu, searchTypeMenu, Uncheck, %search_RemoveZeroes%
    Menu, searchTypeMenu, Uncheck, %search_Fabrimport%
    Menu, searchTypeMenu, Uncheck, %search_Faroluz%
    Menu, searchTypeMenu, Uncheck, %search_Ferrolux%
    Menu, searchTypeMenu, Uncheck, %search_Solnic%
    Menu, searchTypeMenu, Check, %type%
	
	if(!initial)
	{
		if(type != search_Default)
		{
			RegWrite, REG_SZ, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSearchType, %type%		
		}
		else
		{
			RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSearchType
		}
	}
}

; Create a submenu in the first menu (a right-arrow indicator). When the user selects it, the second menu is displayed.
Menu, Tray, Add, Search Type, :searchTypeMenu

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
	ControlSetText, %controlName%,, %windowName%
	SendRaw, % newText
	Sleep, 200
}

DeepCopyControl(controlName, windowName1, windowName2)
{
	ControlGet, controlHwnd, Hwnd,,%controlName%,%windowName1%
	controlType := GetClassName(controlHwnd)
	if(controlType == "Edit")
	{
		ControlGetText, controlText, %controlName%, %windowName1%
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
RegRead, savedSearchType, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSearchType
RegRead, savedSuppressWarnings, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSuppressWarnings
RegRead, savedSkipOnSearch, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSkipOnSearch
if(savedModificadores or savedPostSearchString or savedSearchType)
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
	if(savedSearchType)
	{
		explanationConfig = %explanationConfig%`nSearch Type: %savedSearchType%
	}
	if(savedSuppressWarnings)
	{
		explanationConfig = %explanationConfig%`nSuppress Warnings: %savedSuppressWarnings%
	}
	if(savedSkipOnSearch)
	{
		explanationConfig = %explanationConfig%`nSkip on Unsuccessful Search: %savedSkipOnSearch%
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
		if(savedSearchType)
		{
			setSearchType(savedSearchType)
		}
		if(savedSuppressWarnings)
		{
			toggleSuppressWarnings()
		}
		if(savedSkipOnSearch)
		{
			toggleAutoPilot()
		}
	}
	IfMsgBox, Cancel
	{
		RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedModificadores
		RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedPostSearchString
		RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSearchType
		RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSuppressWarnings
		RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSkipOnSearch
	}
}
;}

;{ Keybinds
Launch_Media::
;FastCorrectNota("Monteluz")
;FastSetRubro("16")
;FastSetAlias()
;SetMargins("60", "25", "40", true)
;Msgbox, Testing...	
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
AnteriorArticulo(false)
Buscar()
return

Media_Next::
ProximoArticulo(false)
Buscar()
return

Launch_Mail::
If(WinExist("Diferencia"))
{
	Send, {Left}{Enter}
	return
}
if WinExist(ventFactProov)
{
	if (WinExist("Modificación"))
	{
		ControlGet, tipoIVA, Choice, , ComboBox1, Modificación
		if (InStr(tipoIVA, "NETO GRAVADO"))
		{
			ControlGetText, netoPrevio, Edit1, Modificación
			ControlGetText, netoReal, Static12, %ventFactProov%
			InputBox, totalTemp, Importe Sumado, Ingrese el importe sumado que aparece abajo a la derecha`, al lado del botón Asigna.
			if(ErrorLevel or not IsNum(totalTemp))
			{
				return
			}
			newNeto := TextPrice2Float(netoPrevio) - TextPrice2Float(totalTemp) + TextPrice2Float(netoReal)
			ControlFocus, Edit1, Modificación
			ControlSetText, Edit1,, Modificación
			SendRaw, % newNeto
			return
		}
	}

	if (WinExist("Nuevo") or WinExist("Modificación"))
	{
		ControlSend, Edit4, {Enter}, Nuevo
		ControlSend, Edit4, {Enter}, Modificación
		WinWait, %ventModificarArticulo%
		ControlGetText, factPrecioCosto, %campoPrecioCosto%, %ventModificarArticulo%
		factPrecioCosto := TextPrice2Float(factPrecioCosto)
		ApplyPriceMultipliers(factPrecioCosto)
		
		ControlClick, %botonProovedoresHabituales%, %ventModificarArticulo%,,,, NA ;Clickea el boton Proveedores Habituales
		WinWait, %ventProveedoresHabituales%, , 5
		if ErrorLevel {
			MsgBox, GetAlias - Could not rouse ventProveedoresHabituales from the dead.
			return
		}
		ControlClick, Ver, %ventProveedoresHabituales%,,,, NA
		WinWait, %ventVerProveedor%, , 5
		if ErrorLevel
		{
			MsgBox, GetAlias - Could not summon ventVerProveedor to this mortal coil.
			return
		}
		ControlGetText, factAliasText, %campoAlias_Habituales%, %ventVerProveedor%
		ControlSend,, {Esc}, %ventVerProveedor%
		ControlSend,, {Esc}, %ventProveedoresHabituales%
		ControlSend, Cancela, {Space}, %ventModificarArticulo%
		ControlFocus, TWBROWSE1, %ventFactProov%
		ControlSend, TWBROWSE1, {PGDN}, %ventFactProov%
		
		finalDetailText = %factAliasText% - %factPrecioCosto%`r`n
		LogSend(finalDetailText)
		Sleep, 200
		ControlClick, Button1, %ventFactProov%,,,, NA

	}
	return
}

WinWait, %ventModificarArticulo%
ControlSend, Ok, {Space}, %ventModificarArticulo%
ControlSend, Sí, {Enter}, Atención
ControlSend, Sí, {Enter}, Atención
ProximoArticulo(false)
Buscar()
return

!^Launch_Mail::
camposAClonar := ["Edit3", "Edit5", "Edit6", "ComboBox1", "ComboBox5", "Edit8", "Edit11", "Edit14", "ComboBox2", "ComboBox3", "ComboBox4"]
for i, elCampo in camposAClonar
{
	DeepCopyControl(elCampo, "Modifica", "Nuevo")
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
If(InStr(WindowUnderMouse(), ventReporteArticulos))
{
	WinKill, %ventModificarArticulo%
}
Else If(InStr(WindowUnderMouse(), ventAdobeReader) or InStr(WindowUnderMouse(), ventCalc) or InStr(WindowUnderMouse(), ventWord))
{
	Click, 2
	Sleep, 300
	WinWait, A
	Send {Ctrl Down}c{Ctrl Up}
	Sleep, 100
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

RemoveToolTip:
ToolTip
return

Exit:
ExitApp
return