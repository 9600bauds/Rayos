#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance force ; Close old versions of this script automatically.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2 ; Match window titles anywhere, not just at the start.

;{ Globals (most of these are effectively defines)
global ventModificarArticulo := "Modificación : "
global campoCodigo := "Edit1"
global campoPrecioCosto := "Edit8"
global campoNota := "Edit18"

global ventModificarProveedor := "Modificación"
global ventProveedoresHabituales := "Proveedores Habituales"
global campoAlias_Habituales := "Edit2"
global ventAliasProveedor := "Alias del Proveedor"
global campoAlias_Dedicado := "Edit1"

global ventReporteArticulos := "Artículos a Modificar"
global ventListaArticulos := "ARTICULOS-LA CASA DEL ELECTRICISTA"
global campoListado := "TXBROWSE1"

global ventNotepad := "ahk_class Notepad"
global ventWord := "ahk_exe WINWORD.EXE"
global ventCalc := "OpenOffice Calc"
global ventCalc_Buscar := "Find & Replace"
global ventCalc_Main := "ahk_class SALFRAME" ;Precisamente la planilla principal, no ningún diálogo
global ventAdobeReader := "Adobe Acrobat Reader"
global ventAdobeReader_Buscar := "ahk_class AVL_AVWindow"
global ventAdobeReader_BuscarOK := "Button18"
global ventAdobeReader_Buscar_Input := "Edit5"
global ventAbodeReader_Buscar_Matches := "Static12"


global search_Default = "Default"
global search_Exact = "Exact"
global search_Start = "Match Start"
global search_End = "Match End"
global search_RemoveLastWord = "Remove Last Word"
global search_LongestNumber = "Longest Number"
global search_Fabrimport = "Fabrimport"
global search_Faroluz = "Faroluz"
global search_Ferrolux = "Ferrolux"
global search_Solnic = "Solnic"
global searchType := "Default"

global suppressWarnings := false
global autoPilot := true
global overrideMiddleClick := true

global modificadoresText := "+0%" ;These two should be equivalent and are only set once in SetModificadores().
global modificadoresMult := 1

global lastPercent := 0
global preciosGuardados := {}

global PostSearchString := ""
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
	
	modificadoresText := tempStr
	modificadoresMult := tempTally
	tempPercent := Multiplier2Percent(modificadoresMult)
	RegWrite, REG_SZ, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedModificadores, %modificadoresText%
	if(displayMessage)
	{
		MsgBox, Modificadores actualizados. Nuevos modificadores:`n%modificadoresText%`n(%tempPercent% / x%modificadoresMult%)
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
	else if(searchType == search_RemoveLastWord){
		alias := RegExReplace(alias, " \w+$", "")
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
	else if(searchType == search_Fabrimport){
		alias := "[^0-9]" . alias . "$"
	}
	else if(searchType == search_Faroluz){
		alias := RegExReplace(alias, " \w+$", "") . "$"
	}
	else if(searchType == search_Ferrolux){
		RegExMatch(alias, "([A-Z]+-\d+)", alias)
	}
	else if(searchType == search_Solnic){
		alias := "^" . alias . "[\s+|$]"
	}
	
	return alias
}

GetAlias(parseAfter := true, checkNota := true){
	aliasText := ""
	if(WinExist(ventReporteArticulos))
	{
		;ControlClick, TBTNBMP29, %ventReporteArticulos% ;silvina proofing
		;WinWait, %ventAliasProveedor% ;silvina proofing
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
				ControlClick, Modifica, %ventListaArticulos%,,,, NA
				ControlClick, Modifica, %ventReporteArticulos%,,,, NA
				ControlSend, Modifica, {Space}, %ventListaArticulos%
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
				ControlClick, Button7, %ventModificarArticulo%,,,, NA ;Clickea el boton Proveedores Habituales
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
			ControlClick, Modifica, %ventListaArticulos%,,,, NA
			ControlClick, Modifica, %ventReporteArticulos%,,,, NA
			ControlSend, Modifica, {Space}, %ventListaArticulos%
			ControlSend, Modifica, {Space}, %ventReporteArticulos%
			WinWait, %ventModificarArticulo%, , 5
			if ErrorLevel {
				MsgBox, GetAlias - Could not cometh here ventModificarArticulo.
				return
			}
		}
		ControlGetText, notaAdicional, %campoNota%, %ventModificarArticulo%
		RegExMatch(notaAdicional, ".*[Alias completo|Alias|Simil]: (.*)$", aliasReplacement)
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
	{ ;TODO TEST
        WinActivate, %ventAdobeReader_Buscar%
        WinWait, %ventAdobeReader_Buscar%
        ControlClick, %ventAdobeReader_BuscarOK%, %ventAdobeReader_Buscar%
        WinWait, %ventAdobeReader_Buscar%
        ControlFocus, %ventAdobeReader_Buscar_Input%, %ventAdobeReader_Buscar%
		SendRaw, % alias
		Send, {Enter}
		
		WinWait, %ventAdobeReader_Buscar%
		WaitControlNotExist("Stop", ventAdobeReader_Buscar)
        WinWait, %ventAdobeReader_Buscar%
		
        ControlGetText, resultsText, %ventAbodeReader_Buscar_Matches%, %ventAdobeReader_Buscar%
        if(InStr(resultsText, "0 doc")){
            OnUnsuccessfulSearch()
            return 0
        }
		else if(InStr(resultsText, "1 instance(s)")){
			ControlClick, AVSearchTreeDocItemView, %ventAdobeReader_Buscar%
            OnSuccessfulSearch()
            return 1
        }
        else if(InStr(resultsText, "instance(s)")){
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
	
}

OnUnsuccessfulSearch(){
}
OnSuccessfulSearch(){
    WinActivate, %ventCalc_Main%
	
	if(SubStr(PostSearchString,0,1) == "c"){
		trimmed := SubStr(PostSearchString,1,StrLen(PostSearchString)-1)		
		Send, %trimmed%
		Send {Ctrl Down}c{Ctrl Up}
		WinWait, A
		if(autoPilot) ;living on a edge baby
		{
			success := PastePrice()
			if(success and lastPercent < 20 and lastPercent > -15) ;living on a EEEEDGE
			{
				Send, {Launch_Mail}
			}
		}

	}
	else{
		Send, %PostSearchString%
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

PastePrice(newPrice := 0){
	if(not WinExist(ventModificarArticulo))
	{
		ControlSend, Modifica, {Space}, %ventListaArticulos%
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

	ControlGetText, notaAdicional, %campoNota%, %ventModificarArticulo%

	precioAdicionalString := ""
	RegExMatch(notaAdicional, ".*Incluye (.*)$", preciosAdicionales)
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
		precioAdicionalString := " +" . preciosGuardados[preciosAdicionales1]
		newPrice := newPrice + preciosGuardados[preciosAdicionales1]
	}
	
	modificadorAdicionalString := ""
    RegExMatch(notaAdicional, ".*Precio de lista \*([0-9.]+)", extraMults)
    if(extraMults1)
	{
        modificadorAdicionalString := " *" . extraMults1
        newPrice := newPrice * extraMults1
    }
    RegExMatch(notaAdicional, ".*Precio de lista \/([0-9.]+)", extraDivisions)
    if(extraDivisions1)
	{
		modificadorAdicionalString := " /" . extraDivisions1
        newPrice := newPrice / extraDivisions1
    }
	
	newPrice := newPrice * modificadoresMult
	
	if(newPrice * 500 < oldPrice){ ;FUCK THOUSANDS SEPARATORS
        newPrice := newPrice * 1000
    }

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
;}

;{ Navegación
ProximoArticulo(openAfter := true)
{
	if(WinExist(ventModificarArticulo))
	{
		WinKill, %ventModificarArticulo%
	}
	ControlSend, Sí, {Enter}, Atención
	ControlSend, Sí, {Enter}, Atención
	ControlSend, %campoListado%, {Down}, %ventListaArticulos%
	ControlSend, %campoListado%, {Down}, %ventReporteArticulos%
	if(openAfter)
	{
		ControlClick, Modifica, %ventListaArticulos%,,,, NA
		ControlClick, Modifica, %ventReporteArticulos%,,,, NA
		;ControlSend, Modifica, {Space}, %ventListaArticulos% ;Clickea el boton Modifica
	}
}

AnteriorArticulo(openAfter := true)
{
	if(WinExist(ventModificarArticulo))
	{
		WinKill, %ventModificarArticulo%
	}
	ControlSend, %campoListado%, {Up}, %ventListaArticulos%
	ControlSend, %campoListado%, {Up}, %ventReporteArticulos%
	if(openAfter)
	{
		ControlClick, Modifica, %ventListaArticulos%,,,, NA
		ControlClick, Modifica, %ventReporteArticulos%,,,, NA
		;ControlSend, Modifica, {Space}, %ventListaArticulos% ;Clickea el boton Modifica
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
		explanation := "Write out a set of instructions to send after a successful search.`r`rEach instruction must be between curly brackets, such as: {Right}`rAdd a number after your instruction to make it repeat that many times, for example: {Right 2}.`rSyntax is the same as AutoHotKey's Send command.`rFinish the string with a lowercase c outside brackets to send CTRL+C after. Example: {Right 2}c"
		if(PostSearchString == "")
		{
			defaultInput := "{Right 2}"
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
Menu, searchTypeMenu, Add, %search_RemoveLastWord%, setSearchRemoveLastWord, Radio
setSearchRemoveLastWord(){
    setSearchType(search_RemoveLastWord)
}
Menu, searchTypeMenu, Add, %search_LongestNumber%, setSearchLongestNumber, Radio
setSearchLongestNumber(){
    setSearchType(search_LongestNumber)
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
    Menu, searchTypeMenu, Uncheck, %search_RemoveLastWord%
    Menu, searchTypeMenu, Uncheck, %search_LongestNumber%
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
    }
    else{
        Menu, Tray, Check, Suppress Warnings
        suppressWarnings := true
    }
}

Menu, Tray, Add  ; Add a separator line.
Menu, Tray, Add, Exit, Exit
;}

;{ Misc
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
	oCell := oDoc.getCurrentSelection
	Col:=oCell.CellAddress.Column
	Row:=oCell.CellAddress.Row 
	FinalStr := Col "-" Row
	Return FinalStr
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
	}
	IfMsgBox, Cancel
	{
		RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedModificadores
		RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedPostSearchString
		RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSearchType
	}
}
;}

;{ Keybinds
Launch_Media::
;Msgbox, Testing...
return

^Launch_Media::
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
WinWait, %ventModificarArticulo%
ControlSend, Ok, {Space}, %ventModificarArticulo%
ControlSend, Sí, {Enter}, Atención
ControlSend, Sí, {Enter}, Atención
ProximoArticulo(false)
Buscar()
return

Browser_Search::
Buscar()
return

Browser_Home::
PastePrice()
return

^Browser_Home::
return

#If overrideMiddleClick and not InStr(WindowUnderMouse(), "Firefox")
MButton::
If(InStr(WindowUnderMouse(), ventReporteArticulos))
{
	WinKill, %ventModificarArticulo%
}
Else
{
	Click, 2
	Sleep, 300
	WinWait, A
	Send {Ctrl Down}c{Ctrl Up}
	Sleep, 100
	WinWait, A
	Send {Browser_Home}
}
return
#If

^Esc::
Pause
return

Exit:
ExitApp
return