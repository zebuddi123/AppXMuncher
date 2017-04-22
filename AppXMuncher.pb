; -----------------------------------------------------------------------------
;           Name: AppXMuncher
;    Description: Front -End List`s Windows 10 AppX`s installed for all users and List`s allowing selection for uninstall via a created powershell script.
;         Author: Zebuddi
;           Date: 2017-04-22
;        Version: 0.1
;     PB-Version: 5.6
;             OS: Windows 10
;         Credit:
;          Forum:
; -----------------------------------------------------------------------------

Structure APPXDATA 
	Name.s
	Publisher.s
	Architecture.s
	ResourceId.s
	Version.s
	PackageFullName.s
	InstallLocation.s
	IsFramework.s
	PackageFamilyName.s
	PublisherId.s
	PackageUserInformation.s
	IsResourcePackage.s
	IsBundle.s
	IsDevelopmentMode.s
	Dependencies.s
EndStructure



Global NewList _llAppX.APPXDATA()

Procedure.s sReadFileToString(sFileName.s)
	Protected *buffer, iNBytes.i, iFileID.i, sStringData.s,  bBOM.b
	If ReadFile(iFileID, sFileName) And Lof(iFileID) > #Null
		bBOM =  ReadStringFormat(iFileID)
		*buffer = AllocateMemory(Lof(iFileID))
		If *buffer
			iNBytes =ReadData(iFileID, *buffer,  Lof(iFileID)) 
			sStringData = PeekS(*buffer, iNBytes, bBOM)
			FreeMemory(*buffer)
			CloseFile(iFileID)
			ProcedureReturn sStringData
		EndIf
	EndIf
EndProcedure

Procedure CleanUP()
	FreeList(_llAppX())
EndProcedure

Procedure.s sMakeRandomName(iNameLength)
	Protected iIndex.i, sReturnString.s
	Protected sAlpha.s ="abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRESTUVWXYZ0987654321"
	For iIndex = 1 To iNameLength
		sReturnString + Mid(sAlpha,Random(62,1),1)
	Next
	ProcedureReturn sReturnString + "_pshell.txt"
EndProcedure


Procedure ProcessAppX() 
	Protected iRegexElements.i 			= CreateRegularExpression(#PB_Any, "^Name.*?(?=^Name)", #PB_RegularExpression_DotAll|#PB_RegularExpression_AnyNewLine|#PB_RegularExpression_MultiLine)
	Protected iRegex.i 					= CreateRegularExpression(#PB_Any, "^.+?:\s")
	Protected iRegexDependencies.i		= CreateRegularExpression(#PB_Any, "(?<=Dependencies\s{11}:)\s\{.*\}?", #PB_RegularExpression_DotAll|#PB_RegularExpression_AnyNewLine|#PB_RegularExpression_MultiLine)
	Protected Dim t$(0), Dim dep$(0)
	Protected bProgramRun.b, sFileString.s, iNbr.i, iIndex.i, sParamter.s, tempfile$ = GetTemporaryDirectory()+ sMakeRandomName(10)
	-
	sParamter = "Get-AppxPackage -AllUser > " + tempfile$
	SetClipboardText(sParamter)
	bProgramRun = RunProgram("powershell.exe", sParamter, "", #PB_Program_Hide|#PB_Program_Wait)	
	sFileString.s = sReadFileToString(tempfile$)+ #CRLF$ + "Name"
	
	If sFileString
		iNbr = ExtractRegularExpression(iRegexElements, sFileString, t$())
		
		With _llAppX()
			For iIndex = 0 To iNbr-1
				AddElement(_llAppX())
				\Name 					= Trim(ReplaceRegularExpression(iRegex, StringField(t$(iIndex), 1,  #CRLF$), ""))
				\Publisher 				= Trim(ReplaceRegularExpression(iRegex, StringField(t$(iIndex), 2,  #CRLF$), ""))
				\Architecture 			= Trim(ReplaceRegularExpression(iRegex, StringField(t$(iIndex), 3,  #CRLF$), ""))
				\ResourceId 			= Trim(ReplaceRegularExpression(iRegex, StringField(t$(iIndex), 4,  #CRLF$), ""))
				\Version 				= Trim(ReplaceRegularExpression(iRegex, StringField(t$(iIndex), 5,  #CRLF$), ""))
				\PackageFullName 		= Trim(ReplaceRegularExpression(iRegex, StringField(t$(iIndex), 6,  #CRLF$), ""))
				\InstallLocation 		= Trim(ReplaceRegularExpression(iRegex, StringField(t$(iIndex), 6,  #CRLF$), ""))
				\IsFramework 			= Trim(ReplaceRegularExpression(iRegex, StringField(t$(iIndex), 8,  #CRLF$), ""))
				\PackageFamilyName 		= Trim(ReplaceRegularExpression(iRegex, StringField(t$(iIndex), 9,  #CRLF$), ""))
				\PublisherId 			= Trim(ReplaceRegularExpression(iRegex, StringField(t$(iIndex), 10, #CRLF$), ""))
				\PackageUserInformation	= Trim(ReplaceRegularExpression(iRegex, StringField(t$(iIndex), 11, #CRLF$), ""))
				\IsResourcePackage 		= Trim(ReplaceRegularExpression(iRegex, StringField(t$(iIndex), 12, #CRLF$), ""))
				\IsBundle 				= Trim(ReplaceRegularExpression(iRegex, StringField(t$(iIndex), 13, #CRLF$), ""))
				\IsDevelopmentMode 		= Trim(ReplaceRegularExpression(iRegex, StringField(t$(iIndex), 14, #CRLF$), ""))
				ExtractRegularExpression(iRegexDependencies, t$(iIndex), dep$()): \Dependencies		 	= Trim(dep$(0)) 
			Next
		EndWith
		
	EndIf
	FreeArray(dep$())
	FreeArray(t$())
	FreeRegularExpression(iRegex)
	FreeRegularExpression(iRegexDependencies)
	FreeRegularExpression(iRegexElements)
EndProcedure

;---- Main -------

ProcessAppX()
; CleanUP()

CallDebugger


