Option Explicit

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005

Function LoadRegUninstall(ByVal fHKEY, ByVal strBaseKey)
	Dim objReg
	Dim strReturn
	Dim strSubKey, arrSubKeys, strValue, intRet
	Dim strDisplayName, strDisplayVersion, strInstallDate, strInstallLocation, strUninstallString, strQuietUninstallString, strURLInfoAbout

	Set objReg = GetObject("WinMgmts:\\.\Root\default:StdRegProv")

	strReturn = ""

	objReg.EnumKey fHKEY, strBaseKey, arrSubKeys

	If (IsArray(arrSubKeys) = -1) Then
		strReturn = strReturn & "DisplayName;DisplayVersion;InstallDate;InstallLocation;UninstallString;DisplayVersion;QuietUninstallString;URLInfoAbout" & VbCrLf

		For Each strSubKey In arrSubKeys
			intRet = objReg.GetStringValue(fHKEY, strBaseKey & strSubKey, "DisplayName", strValue)
			
			If intRet <> 0 Then 
				intRet = objReg.GetStringValue(fHKEY, strBaseKey & strSubKey, "QuietDisplayName", strValue)
			End If
			
			strDisplayName = strValue

			If (strValue <> "") and (intRet = 0) Then
				objReg.GetStringValue fHKEY, strBaseKey & strSubKey, "DisplayVersion", strDisplayVersion
				objReg.GetStringValue fHKEY, strBaseKey & strSubKey, "InstallDate", strInstallDate
				objReg.GetStringValue fHKEY, strBaseKey & strSubKey, "InstallLocation", strInstallLocation
				objReg.GetStringValue fHKEY, strBaseKey & strSubKey, "UninstallString", strUninstallString
				objReg.GetStringValue fHKEY, strBaseKey & strSubKey, "QuietUninstallString", strQuietUninstallString
				objReg.GetStringValue fHKEY, strBaseKey & strSubKey, "URLInfoAbout", strURLInfoAbout
				
				strReturn = strReturn & strDisplayName & ";" & strDisplayVersion & ";" & strInstallDate & ";" & strInstallLocation & ";" & strUninstallString & ";" & strDisplayVersion & ";" & strQuietUninstallString & ";" & strURLInfoAbout & VbCrLf 
			End If
		Next
	End If

	Set objReg = Nothing

	LoadRegUninstall = strReturn
End Function

'LoadRegUninstall HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\", "C:\temp\InstalledSoft.csv"
'LoadRegUninstall HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Uninstall\", "C:\temp\InstalledSoft.csv"
'LoadRegUninstall HKEY_LOCAL_MACHINE, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\", "C:\temp\InstalledSoft.csv"