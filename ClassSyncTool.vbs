' Версия 07-09-2020
Option Explicit

Class SyncTool

	Private objWshShell
	Private objFSO
	Private objSysInfo
	Private wshEnviromentUser
	Private wshEnviromentProcess
	Private objSWbemServices
	Private objLDAPUser

	Private dictEnv

	Private strLogFile
	Private strScrName
	Private strJobName

	Private objLogFile

	Private Sub Class_Initialize
		'On Error Resume Next

		Set objWshShell = WScript.CreateObject("WScript.Shell")
		Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
		Set wshEnviromentUser = objWshShell.Environment("User")
		Set wshEnviromentProcess = objWshShell.Environment("Process")
		Set objSWbemServices = GetObject("winmgmts:\\.\root\cimv2")
		Set objSysInfo = CreateObject("ADSystemInfo")
		' Информация из домена
		'Do
		'	Sleep 5
		'Loop While (wshEnviromentProcess("LOGONSERVER") = "") ' Если домен не доступен, попробовать еще раз
		'Set objLDAPUser = GetObject("LDAP://fss.local:389/" & objSysInfo.UserName)
		Set dictEnv = CreateObject("Scripting.Dictionary")
		
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

	Public Default Function Init(job)
		'On Error Resume Next

		Dim strItem, intEq, strName, strValue

		For Each strItem In wshEnviromentUser
			'If Left(strItem, 3) = "FSS" Then
				intEq = InStr(1, strItem, "=" , vbTextCompare)
				strName = Left(strItem, intEq - 1)
				strValue = Mid(strItem, intEq + 1)
				dictEnv.Add strName, strValue
			'End If
		Next

		dictEnv.Add "UserName", objWshShell.ExpandEnvironmentStrings("%USERNAME%")
		dictEnv.Add "ComputerName", objWshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
		dictEnv.Add "AppData", objWshShell.ExpandEnvironmentStrings("%APPDATA%") & "\" & wshEnviromentUser("FSSFolderName")

		strJobName = job

		strLogFile = objWshShell.ExpandEnvironmentStrings("%APPDATA%") & "\" & wshEnviromentUser("FSSFolderName") & "\" & "sync.log"
		strScrName = WScript.ScriptName & "[" & strJobName & "]"

		If Err.Number <> 0 Then DebugError(Err)

		Set Init = Me
	End Function

	Public Function Env(ByVal strName)
		Env = dictEnv.Item(strName)
	End Function

	Public Property Get JobName()
		JobName = strJobName
	End Property

	Public Property Let PathToLogFile(strValue)
		strLogFile = strValue
	End Property

	Public Property Get PathToLogFile()
		PathToLogFile = strLogFile
	End Property

	Public Sub SetEnviroment(ByVal strName, ByVal strValue)
		Debug "Установка переменной окружения [User]" & Quotes(strName) & " = " & Quotes(strValue)
		wshEnviromentUser(strName) = strValue
	End Sub

	Public Function GetEnviroment(ByVal strName)
		GetEnviroment = wshEnviromentUser(strName)
	End Function

	Public Sub SetEnviromentProcess(ByVal strName, ByVal strValue)
		wshEnviromentProcess(strName) = strValue
	End Sub

	Public Function GetEnviromentProcess(ByVal strName)
		GetEnviromentProcess = wshEnviromentProcess(strName)
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
		Loop While (Err.Number <> 0) and (i < 10) ' Если фаил занят, попробовать еще несколько раз
	End Sub

	Public Sub DebugError(objErr)
		If objErr.Number <> 0 Then
			Debug "Ошибка: (" & objErr.Number & ") " & objErr.Description
			objErr.Clear
		End If
	End Sub

	' 
	Public Sub DebugClear()
		DebugClearIfSizeByName strLogFile, 0
	End Sub

	' Удалять лог фаил при превышении intSize килобайт 
	Public Sub DebugClearIfSize(ByVal intSize)
		DebugClearIfSizeByName strLogFile, intSize
	End Sub

	' Удалять лог фаил при превышении intSize килобайт, с указанием пути к файлу
	Public Sub DebugClearIfSizeByName(ByVal strName, ByVal intSize)
		On Error Resume Next
		Set objLogFile = objFSO.GetFile(strName)
		If objLogFile.Size > (intSize * 1000) Then
			objFSO.DeleteFile strName, True
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

		If objFSO.FolderExists(strDir) Then
			Debug "Папка " & Quotes(strDirName) & " уже существует."
		Else
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
					Debug "Создание папки " & Quotes(strDirName)
					objFSO.CreateFolder strDirBuild
				End if
			Next
		End If

		DebugError(Err)
	End Sub

	Public Sub CopyFolder(ByVal strSource, ByVal strDestination, ByVal bRewrite)
		CreateDirs strDestination
		Debug "Копирование из " & Quotes(strSource) & " в " & Quotes(strDestination)
		objFSO.CopyFolder strSource, strDestination, True
		Debug "Копирование завершено"
	End Sub

	Public Sub CopyFile(ByVal strSource, ByVal strDestination, ByVal bRewrite)
		Debug "Копирование из " & Quotes(strSource) & " в " & Quotes(strDestination)
		objFSO.CopyFile strSource, strDestination, bRewrite
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
			'Debug "Завершено."
		End If
		
		Set objWshExec = Nothing
	End Sub

	Public Sub SendMsgAdmin(ByVal strText)
		On Error Resume Next

		Dim objHttp, res

		Set objHttp = CreateObject("Msxml2.ServerXMLHTTP")
		
		Debug "Сообщение для админа: " & strText
		objHttp.Open "GET", "http://xmpp.6604.local/sendmessage_example.php?msg=" & strText, False, "server.6604", "server66"
		objHttp.Send
		
		res = objHttp.ResponseText
		'Debug "Результат отправки: " & res

		Set objHttp = Nothing
	End Sub

	Public Function CountProcess(ByVal strProcessName)
		Dim colSWbemObjectSet

		Set colSWbemObjectSet = objSWbemServices.ExecQuery("SELECT * FROM Win32_Process where Name = '" & strProcessName & "'")
		CountProcess = colSWbemObjectSet.Count

		Set colSWbemObjectSet = Nothing
	End Function

	Public Function GetLDAP(ByVal strValue)
		Set objLDAPUser = GetObject("LDAP://fss.local:389/" & objSysInfo.UserName)
		GetLDAP = objLDAPUser.Get(strValue)
	End Function

	Public Function DiskSpaceFree(ByVal strValue)
		Dim disks, gb, e, free_space_gb

		Set disks = objSWbemServices.ExecQuery("select * from Win32_LogicalDisk where DriveType=3")

		free_space_gb = -1
		gb = 1024*1024*1024

		For Each e in disks
			'Debug e.DeviceID
			if e.DeviceID = strValue Then
				free_space_gb=e.FreeSpace/gb
			End If
		Next

		disks = Nothing

		DiskSpaceFree = free_space_gb

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

	Public Function FileExists(ByVal strValue)
		FileExists = objFSO.FileExists(strValue)
	End Function

	Public Function GetFileVersion(ByVal strValue)
		GetFileVersion = objFSO.GetFileVersion(strValue)
	End Function

End Class