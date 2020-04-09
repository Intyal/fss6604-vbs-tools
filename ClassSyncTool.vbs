' Версия 12-03-2020
Option Explicit

Class SyncTool

	Private objWshShell
	Private objFSO
	Private objSysInfo
	Private wshEnviromentUser
	Private objSWbemServices
	Private objLDAPUser

	Private strMyFolderName
	Private strMyPath
	Private strMyPathNet
	Private strMyPathPublic
	Private strMyGeneralGroup
	
	Private strLoginName
	Private strComputerName
	Private strAppDataFolder
	Private strUserFullName
	Private strLDAPUserName

	Private strLogFile
	Private strScrName

	Private objLogFile

	Private Sub Class_Initialize
		On Error Resume Next

		Set objWshShell = WScript.CreateObject("WScript.Shell")
		Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
		Set objSysInfo = CreateObject("ADSystemInfo")
		Set wshEnviromentUser = objWshShell.Environment("User")
		Set objSWbemServices = GetObject("winmgmts:\\.\root\cimv2")
		Set objLDAPUser = GetObject("LDAP://" & objSysInfo.UserName)
	End Sub

	Private Sub Class_Terminate
		Set objWshShell = Nothing
		Set objFSO = Nothing
		Set objSysInfo = Nothing
		Set wshEnviromentUser = Nothing
		Set objSWbemServices = Nothing
		Set objLDAPUser = Nothing
	End Sub

	' -------------------------------------------------------------------------

	Public Default Function Init(strJobName)
		strMyPath = wshEnviromentUser("FSSPath")
		strMyPathNet = wshEnviromentUser("FSSPathNet")
		strMyPathPublic = wshEnviromentUser("FSSPathPublic")
		strMyGeneralGroup = wshEnviromentUser("FSSGeneralGroup")
		strMyFolderName = wshEnviromentUser("FSSFolderName")

		strLoginName = objWshShell.ExpandEnvironmentStrings("%USERNAME%")
		strComputerName = objWshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
		strAppDataFolder = objWshShell.ExpandEnvironmentStrings("%APPDATA%") & "\" & strMyFolderName

		Dim arrName
		strLDAPUserName = objSysInfo.UserName
		arrName = split(objSysInfo.UserName, ",")
		strUserFullName = mid(arrName(0), 4)

		strLogFile = strAppDataFolder & "\sync.log"
		strScrName = WScript.ScriptName & "[" & strJobName & "]"

		'Debug "ClassSyncTools init"
		'Debug "strMyPath = " & strMyPath
		'Debug "strMyPathNet = " & strMyPathNet
		'Debug "strMyPathPublic = " & strMyPathPublic
		'Debug "strMyGeneralGroup = " & strMyGeneralGroup
		'Debug "strMyFolderName = " & strMyFolderName
		'Debug "strLoginName = " & strLoginName
		'Debug "strComputerName = " & strComputerName
		'Debug "strAppDataFolder = " & strAppDataFolder
		'Debug "strUserFullName = " & strUserFullName
		'Debug "strLogFile = " & strLogFile
		'Debug "strScrName = " & strScrName

		Set Init = Me
	End Function

	Public Property Let PathToLogFile(strValue)
		strLogFile = strValue
	End Property

	Public Property Get PathToLogFile()
		PathToLogFile = strLogFile
	End Property

	Public Property Let SyncFolderName(strValue)
		strMyFolderName = strValue
		SetEnviroment "FSSFolderName", strValue
	End Property

	Public Property Get SyncFolderName()
		SyncFolderName = strMyFolderName
	End Property

	Public Property Let LocalPathToSync(strValue)
		strMyPath = strValue
		SetEnviroment "FSSPath", strValue
	End Property

	Public Property Get LocalPathToSync()
		LocalPathToSync = strMyPath
	End Property

	Public Property Let NetPathToSync(strValue)
		strMyPathNet = strValue
		SetEnviroment "FSSPathNet", strValue
	End Property

	Public Property Get NetPathToSync()
		NetPathToSync = strMyPathNet
	End Property

	Public Property Let PathToPublic(strValue)
		strMyPathPublic = strValue
		SetEnviroment "FSSPathPublic", strValue
	End Property

	Public Property Get PathToPublic()
		PathToPublic = strMyPathPublic
	End Property

	Public Property Let GeneralGroup(strValue)
		strMyGeneralGroup = strValue
		SetEnviroment "FSSGeneralGroup", strValue
	End Property

	Public Property Get GeneralGroup()
		GeneralGroup = strMyGeneralGroup
	End Property

	Public Property Get LoginName()
		LoginName = strLoginName
	End Property

	Public Property Get ComputerName()
		ComputerName = strComputerName
	End Property

	Public Property Get AppDataFolder()
		AppDataFolder = strAppDataFolder
	End Property

	Public Property Get UserFullName()
		UserFullName = strUserFullName
	End Property

	Public Property Get LDAPUserName()
		LDAPUserName = strLDAPUserName
	End Property

	Public Sub SetEnviroment(ByVal strName, ByVal strValue)
		wshEnviromentUser(strName) = strValue
	End Sub

	Public Function GetEnviroment(ByVal strName)
		GetEnviroment = wshEnviromentUser(strName)
	End Function

	Public Function Quotes(ByVal strValue)
		Quotes = """" & strValue & """"
	End Function

	Public Function Recode(StrText, SrcCode, DestCode)
		With CreateObject("ADODB.Stream")
			.Type = 2
			.Mode = 3
			.Charset = DestCode
			.Open
			.WriteText (strText)
			.Position = 0
			.Charset = SrcCode
			Recode = .ReadText
			.Close
		end with
	End Function

	Public Sub Debug(ByVal strText)
		On Error Resume Next
		Dim i

		i = 0
		Do
			Err.Clear
			Set objLogFile = objFSO.OpenTextFile(strLogFile, 8, True)
			objLogFile.WriteLine Now & " " & strScrName & ": " & strText
			objLogFile.Close
			Set objLogFile = Nothing
			If Err.Number <> 0 Then Sleep 1
			i = i + 1
		Loop While (Err.Number <> 0) and (i < 10)
	End Sub

	Public Sub DebugError(objErr)
		If objErr.Number <> 0 Then
			Debug "Ошибка: (" & objErr.Number & ") " & objErr.Description
			objErr.Clear
		End If
	End Sub

	' 
	Public Sub DebugClear()
		On Error Resume Next
		objFSO.DeleteFile strLogFile, True
	End Sub

	' Удалять лог фаил при превышении intSize килобайт 
	Public Sub DebugClearIfSize(ByVal intSize)
		On Error Resume Next
		Set objLogFile = objFSO.GetFile(strLogFile)
		If objLogFile.Size > (intSize * 1000) Then
			objFSO.DeleteFile strLogFile, True
		End If
		Set objLogFile = Nothing
	End Sub

	Public Sub SaveTxt(ByVal strText, ByVal strFile, ByVal bMode)
		Dim objFile, bModeN

		If (strText = "" OR strFile = "") Then Exit Sub

		bModeN = bMode
		If bMode <> 2 Or bMode <> 8 Then
			bModeN = 2
			If bMode = True Then bModeN = 2
			If bMode = False Then bModeN = 8
		End If
		
		Debug "Сохранение в фаил: " & strFile
		Set objFile = objFSO.OpenTextFile(strFile, bModeN, True)
		objFile.WriteLine strText
		objFile.Close
		Set objFile = Nothing

	End Sub

	Public Sub CreateDirs(ByVal strDirName)
		On Error Resume Next

		Dim arrDirs, i, idxFirst, strDir, strDirBuild
		strDir = objFSO.GetAbsolutePathName(strDirName)
		arrDirs = Split(strDir, "\")

		Debug "Создание папки " & strDirBuild
		If Left(strDir, 2) = "\\" Then
			strDirBuild = "\\" & arrDirs(2) & "\" & arrDirs(3) & "\"
			idxFirst    = 4
		Else
			strDirBuild = arrDirs(0) & "\"
			idxFirst    = 1
		End If
		For i = idxFirst to Ubound(arrDirs)
			strDirBuild = objFSO.BuildPath(strDirBuild, arrDirs(i))
			If Not objFSO.FolderExists(strDirBuild) Then
				objFSO.CreateFolder strDirBuild
			End if
		Next

		DebugError(Err)
	End Sub

	Public Sub CopyFolder(ByVal strSource, ByVal strDestination, ByVal bRewrite)
		CreateDirs strDestination
		Debug "Копирование из " & strSource & " в " & strDestination
		objFSO.CopyFolder strSource, strDestination, bRewrite
	End Sub

	Public Sub Sleep(ByVal nSeconds)
		Debug "Пауза " & nSeconds & " сек."
		WScript.Sleep nSeconds * 1000
	End Sub

	Public Sub Run(ByVal strCommand, ByVal intWindowStyle, ByVal bWaitOnReturn)
		' 0 – скрывает окно, будет виден только процесс в диспетчере задач, 1 – нормальный режим, 2 – свернутый вид, 3 – развернутый вид
		If intWindowStyle = "" Then intWindowStyle = 1
		If bWaitOnReturn = "" Then bWaitOnReturn = False
		Debug "Запуск " & strCommand & ", " & intWindowStyle & ", " & bWaitOnReturn
		objWshShell.Run strCommand, intWindowStyle, bWaitOnReturn
	End Sub

	Public Sub RunC(ByVal strCommand)
		Dim objWshExec
		Const WshRunning = 1
		Const WshFailed = 1

		Debug "Выполнение " & strCommand
		Set objWshExec = objWshShell.Exec(strCommand)

		Do While objWshExec.Status = WshRunning
			WScript.Sleep 100
		Loop

		Debug "ProcessID: " & objWshExec.ProcessID
		If objWshExec.ExitCode = WshFailed Then
			Debug "Ошибка: " & objWshExec.StdErr.ReadAll
		Else
			Debug "Результат: " & Recode(objWshExec.StdOut.ReadAll, "cp866", "windows-1251")
		End If
		
		Set objWshExec = Nothing
	End Sub

	Public Function CountProcess(ByVal strProcessName)
		Dim colSWbemObjectSet

		Set colSWbemObjectSet = objSWbemServices.ExecQuery("SELECT * FROM Win32_Process where Name = '" & strProcessName & "'")
		CountProcess = colSWbemObjectSet.Count

		Set colSWbemObjectSet = Nothing
	End Function

	Public Sub PutLDAP(ByVal strName, ByVal strValue)
		Debug "Запись в свойства пользователя " & strName & " = " & strValue
		objLDAPUser.Put strName, strValue
		objLDAPUser.SetInfo
	End Sub

	Public Sub RegWrite(ByVal strKey, ByVal strValue, ByVal strType)
		Debug "Запись в реестр " & strKey & " = " & strValue & "(" & strType & ")"
		objWshShell.RegWrite strKey, strValue, strType
	End Sub

End Class