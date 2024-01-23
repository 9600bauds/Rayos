global searchTypes := []

toggleSearchType(name) {
    for index, searchType in searchTypes {
        if (searchType.name == name) {
            searchType.active := !searchType.active
			break
        }
    }
    
	activeSearchTypesStr := ""
	for index, searchType in searchTypes {
		if (searchType.active)
			activeSearchTypesStr .= searchType.name . "`n"
	}
	; Write the active search types to the registry as REG_MULTI_SZ
	RegWrite, REG_MULTI_SZ, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSearchTypes, %activeSearchTypesStr%
	
	refreshSearchTypeMenu()
}

DeactivateAllSearchTypes() {
    for index_deactsearchtypes, searchType_deactsearchtypes in searchTypes {
        searchType_deactsearchtypes.active := false
    }
    RegDelete, HKEY_CURRENT_USER\SOFTWARE\Rayos, savedSearchTypes
	refreshSearchTypeMenu()
}

refreshSearchTypeMenu() {
    for index, searchType in searchTypes {
        if (searchType.active) {
            Menu, searchTypeMenu, Check, % searchType.name
        } else {
            Menu, searchTypeMenu, Uncheck, % searchType.name
        }
    }
}

createSearchType(name, function) {
    searchType := {name: name, function: function, active: false}
    searchTypes.Push(searchType)
	function := Func("toggleSearchType").Bind(searchType.name)
    Menu, searchTypeMenu, Add, % searchType.name, % function
}

executeSearchTypes(alias) {
    for index, searchType in searchTypes {
        if (searchType.active) {
			functionName := searchType.function
            alias := Func(functionName).Call(alias)
        }
    }
    return alias
}


doSearch_Exact(alias){
    return "^ " . alias . "$"
}
createSearchType("Exact", "doSearch_Exact")

doSearch_Start(alias){
    return "^ " . alias
}
createSearchType("Start", "doSearch_Start")

doSearch_End(alias){
    return alias . "$"
}
createSearchType("End", "doSearch_End")

doSearch_WordBoundaries(alias){
    return "\b" . alias . "\b"
}
createSearchType("Word Boundaries", "doSearch_WordBoundaries")

doSearch_RemoveLastWord(alias){
    return RegExReplace(alias, " \w+$", "")
}
createSearchType("Remove Last Word", "doSearch_RemoveLastWord")

doSearch_RemoveColors(alias){
    alias := RegExReplace(alias, "[A-Za-z]", "")
    return RegExReplace(alias, "\/$", "")
}
createSearchType("Remove Colors", "doSearch_RemoveColors")

doSearch_RemoveLetters(alias){
    return RegExReplace(alias, "[A-Za-z]", "")
}
createSearchType("Remove Letters", "doSearch_RemoveLetters")

doSearch_LongestNumber(alias){
    longestMatch := ""
    for index, match in AllRegexMatches(alias, "[\d]+"){
        if(StrLen(match) > StrLen(longestMatch)){
            longestMatch := match
        }
    }
    return longestMatch
}
createSearchType("Longest Number", "doSearch_LongestNumber")

doSearch_LongestWord(alias){
    longestMatch := ""
    for index, match in AllRegexMatches(alias, "[\w]+"){
        if(StrLen(match) > StrLen(longestMatch)){
            longestMatch := match
        }
    }
    return "\b" . longestMatch . "\b"
}
createSearchType("Longest Word", "doSearch_LongestWord")

doSearch_RemoveZeroes(alias){
    return RegExReplace(alias, "^[0]+", "")
}
createSearchType("Remove Zeroes", "doSearch_RemoveZeroes")

doSearch_Fabrimport(alias){
    return "[^0-9]" . alias . "$"
}
createSearchType("Fabrimport", "doSearch_Fabrimport")

doSearch_Faroluz(alias){
    return RegExReplace(alias, " \w+$", "") . "$"
}
createSearchType("Faroluz", "doSearch_Faroluz")

doSearch_Ferrolux(alias){
    newAlias := alias
    RegExMatch(alias, "([A-Z]+-\d+(\/\d+)?)", newAlias)
    if WinExist(ventCalc)
    {
        return "^ " . newAlias
    } 
    return newAlias
}
createSearchType("Ferrolux", "doSearch_Ferrolux")

doSearch_Solnic(alias){
    return "^" . alias . "[\s+|$]"
}
createSearchType("Solnic", "doSearch_Solnic")

doSearch_WhitespaceOptional(alias){
    return RegExReplace(alias, "\s+", "\s*")
}
createSearchType("Whitespace Optional", "doSearch_WhitespaceOptional")