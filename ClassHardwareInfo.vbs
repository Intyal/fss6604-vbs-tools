Option Explicit
'On Error Resume Next

Class HardwareInfo

	Private objSWbemServices
	Private objQueryResultProcessor
	Private objQueryResultBaseBoard
	Private objQueryResultDesktopMonitor
	Private objQueryResultPhysicalMemory
	Private objQueryResultLogicalDisk
	Private objQueryResultDiskDrive
	Private objQueryResultOperatingSystem

	Private Sub Class_Initialize
		Set objSWbemServices = GetObject("winmgmts:\\.\root\cimv2")
	End Sub

	Private Sub Class_Terminate
		Set objSWbemServices = Nothing
	End Sub

'	Private Sub QuerySplit(objQueryResult, objArray)
'		Dim objItem, i

'		i = 0

'		ReDim objArray(objQueryResult.Count)

'		For Each objItem in objQueryResult
'			Set objArray(i) = objItem
'			i = i + 1
'		Next
'	End Sub

	Private Function TxtTab(ByVal strName, ByVal strValue)
		Dim a, b, c, i
		a = 8
		b = Fix((a * 3 - len(strName)) / 8 + 0.95)
		c = ""
		For i = 1 to b
			c = c & VbTab
		Next
		TxtTab = strName & c & strValue & VbCrLf
	End Function

	' -------------------------------------------------------------------------

	Public Default Function Init()
		'QuerySplit objSWbemServices.ExecQuery("SELECT * FROM Win32_Processor"), objQueryResultProcessor
		'QuerySplit objSWbemServices.ExecQuery("SELECT * FROM Win32_DesktopMonitor"), objQueryResultDesktopMonitor
		'QuerySplit objSWbemServices.ExecQuery("SELECT * FROM Win32_PhysicalMemory"), objQueryResultPhysicalMemory
		Set objQueryResultProcessor = objSWbemServices.ExecQuery("SELECT * FROM Win32_Processor")
		Set objQueryResultBaseBoard = objSWbemServices.ExecQuery("Select * from Win32_BaseBoard")
		Set objQueryResultDesktopMonitor = objSWbemServices.ExecQuery("SELECT * FROM Win32_DesktopMonitor")
		Set objQueryResultPhysicalMemory = objSWbemServices.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
		Set objQueryResultLogicalDisk = objSWbemServices.ExecQuery("SELECT * FROM Win32_LogicalDisk where DriveType=3")
		Set objQueryResultDiskDrive = objSWbemServices.ExecQuery("SELECT * FROM Win32_DiskDrive")
		Set objQueryResultOperatingSystem = objSWbemServices.ExecQuery("SELECT * FROM Win32_OperatingSystem")

		Set Init = Me
	End Function

'	Public Function Processor(ByVal intID, ByVal strValue)
'		Processor = objQueryResultProcessor(intID).Properties_(strValue)
'	End Function

'	Public Function Monitor(ByVal intID, ByVal strValue)
'		Monitor = objQueryResultDesktopMonitor(intID).Properties_(strValue)
'	End Function

'	Public Function Memory(ByVal intID, ByVal strValue)
'		Memory = objQueryResultPhysicalMemory(intID).Properties_(strValue)
'	End Function

	Public Function GetTxt()
		Dim strText, objItem, i, gb, str

		strText = ""
		gb = 1024*1024*1024

		' ������ �� �����������
		i = 1
		strText = strText & VbCrLf & "============================== " & "����������" & " ============================== " & VbCrLf
		For Each objItem In objQueryResultProcessor
			strText = strText & "- " & "��������� � " & i & " -" & VbCrLf
			str = ""
			If objItem.Architecture = 0 Then
				str = "32 ���"
			ElseIf objItem.Architecture = 9 Then
				str = "64 ���"
			End If
			strText = strText & TxtTab("�����������:", str)
			strText = strText & TxtTab("���:", objItem.Name)
			strText = strText & TxtTab("��������:", objItem.Caption)
			strText = strText & TxtTab("���������� ����:", objItem.NumberOfCores)
			strText = strText & TxtTab("���������� ����:", objItem.NumberOfLogicalProcessors)
			strText = strText & TxtTab("������������ ��������:", objItem.MaxClockSpeed & " ���")

			i = i + 1
		Next
		' ������ �� ����������� �����
		i = 1
		strText = strText & VbCrLf & "============================== " & "��������� �����" & " ============================== " & VbCrLf
		For Each objItem In objQueryResultBaseBoard
			strText = strText & TxtTab("�������������:", objItem.Manufacturer)
			strText = strText & TxtTab("������:", objItem.Model)
			strText = strText & TxtTab("�������:", objItem.Product)
			strText = strText & TxtTab("������:", objItem.Version)

			i = i + 1
		Next
		' ������ �� ���������
		i = 1
		strText = strText & VbCrLf & "============================== " & "��������" & " ============================== " & VbCrLf
		For Each objItem In objQueryResultDesktopMonitor
			strText = strText & "- " & "������� � " & i & " -" & VbCrLf
			strText = strText & TxtTab("��������:", objItem.Name)
			strText = strText & TxtTab("����� �� ���� �� X:", objItem.PixelsPerXLogicalInch)
			strText = strText & TxtTab("����� �� ���� �� Y:", objItem.PixelsPerYLogicalInch)
			strText = strText & TxtTab("����� �� ������:", objItem.ScreenWidth)
			strText = strText & TxtTab("����� �� ������:", objItem.ScreenHeight)

			i = i + 1
		Next
		' ������ �� ���
		i = 1
		strText = strText & VbCrLf & "============================== " & "���" & " ============================== " & VbCrLf
		For Each objItem In objQueryResultPhysicalMemory
			strText = strText & "- " & "������ � " & i & " -" & VbCrLf
			strText = strText & TxtTab("�����:", objItem.Capacity / gb & " ��")
			strText = strText & TxtTab("���:", objItem.MemoryType)
			strText = strText & TxtTab("�������� �����:", objItem.SerialNumber)
			strText = strText & TxtTab("����� ���������:", objItem.DeviceLocator)
			strText = strText & TxtTab("��������:", objItem.Speed & " ���")

			i = i + 1
		Next
		' ������ �� ������
		i = 1
		Dim size_gb, free_space_gb, free_percentage
		strText = strText & VbCrLf & "============================== " & "�����" & " ============================== " & VbCrLf
		For Each objItem In objQueryResultLogicalDisk
			strText = strText & "- " & "���������� ���� � " & i & " -" & VbCrLf
			size_gb = objItem.Size / gb
			free_space_gb = objItem.FreeSpace / gb
			free_percentage = free_space_gb * 100 / size_gb
			strText = strText & TxtTab("����� �����:", objItem.DeviceID)
			strText = strText & TxtTab("������ �����:", Fix(size_gb) & " �� ")
			strText = strText & TxtTab("��������:", Fix(free_space_gb) & " ��" & " (" & Fix(free_percentage) & "%)")
			strText = strText & TxtTab("�������� �������:", objItem.FileSystem)

			i = i + 1
		Next
		i = 1
		For Each objItem In objQueryResultDiskDrive
			strText = strText & "- " & "���������� ���� � " & i & " -" & VbCrLf
			size_gb = objItem.Size / gb
			strText = strText & TxtTab("�� ����������:", objItem.DeviceID)
			strText = strText & TxtTab("������ �����:", Fix(size_gb) & "�� ")
			'strText = strText & TxtTab("�������� �����:", objItem.SerialNumber)
			strText = strText & TxtTab("������:", objItem.Status)
			strText = strText & TxtTab("��� ��������� ������:", objItem.LastErrorCode)
			
			i = i + 1
		Next
		' ������ �� ��
		i = 1
		strText = strText & VbCrLf & "============================== " & "��" & " ============================== " & VbCrLf
		For Each objItem In objQueryResultOperatingSystem
			strText = strText & TxtTab("��������:", objItem.Caption)
			strText = strText & TxtTab("������:", objItem.Version)
			strText = strText & TxtTab("������-���:", objItem.ServicePackMajorVersion)
			'strText = strText & TxtTab("�����������:", objItem.OSArchitecture)
			strText = strText & TxtTab("BootDevice:", objItem.BootDevice)
			strText = strText & TxtTab("SystemDevice:", objItem.SystemDevice)
			strText = strText & TxtTab("SystemDirectory:", objItem.SystemDirectory)
			strText = strText & TxtTab("SerialNumber:", objItem.SerialNumber)

			i = i + 1
		Next

		GetTxt = strText
	End Function

	Public Function GetCSV()
		Dim strHeader, strText, objItem, i, gb, str

		strHeader = ""
		strText = ""
		gb = 1024*1024*1024

		' ������ �� �����������
		i = 1
		strHeader = strHeader & "���������(���);���������(�����������);���������(�������);���������(����������, ��);"
		For Each objItem In objQueryResultProcessor
			str = ""
			If i = 1 Then
				If objItem.Architecture = 0 Then
					str = "32"
				ElseIf objItem.Architecture = 9 Then
					str = "64"
				End If
				strText = strText & objItem.Name & ";"
				strText = strText & str & ";"
				strText = strText & objItem.NumberOfLogicalProcessors & ";"
			End If

			i = i + 1
		Next
		strText = strText & (i - 1) & ";"
		' ������ �� ����������� �����
		i = 1
		strHeader = strHeader & "��������� �����(�������������);��������� �����(�������);��������� �����(������);"
		For Each objItem In objQueryResultBaseBoard
			If i = 1 Then
				strText = strText & objItem.Manufacturer & ";"
				strText = strText & objItem.Product & ";"
				strText = strText & "'" & objItem.Version & "'" & ";"
			End If

			i = i + 1
		Next
		' ������ �� ���������
		i = 1
		strHeader = strHeader & "�������(���������� �� X);�������(���������� �� Y);�������(����� �� ����);"
		For Each objItem In objQueryResultDesktopMonitor
			If i = 1 Then
				strText = strText & objItem.ScreenWidth & ";"
				strText = strText & objItem.ScreenHeight & ";"
				strText = strText & objItem.PixelsPerXLogicalInch & ";"
			End If

			i = i + 1
		Next
		' ������ �� ���
		i = 1
		strHeader = strHeader & "���(�����, ��);���(����������, ��);"
		str = 0
		For Each objItem In objQueryResultPhysicalMemory
			str = str + objItem.Capacity / gb

			i = i + 1
		Next
		strText = strText & str & ";"
		strText = strText & (i - 1) & ";"
		' ������ �� ������
		i = 1
		Dim size_gb, free_space_gb, free_percentage
		strHeader = strHeader & "�����([���:�����, ��/��������, ��(%)/������]);"
		For Each objItem In objQueryResultLogicalDisk
			size_gb = objItem.Size / gb
			free_space_gb = objItem.FreeSpace / gb
			free_percentage = free_space_gb * 100 / size_gb
			strText = strText & "[" & objItem.DeviceID & "/" & Fix(size_gb) & "/" & Fix(free_space_gb) & "(" & Fix(free_percentage) & "%)/" & objItem.FileSystem & "]" & ";"

			i = i + 1
		Next
		' ������ �� ��
		i = 1
		strHeader = strHeader & "��(��������);��(������);��(��� ����������)"
		For Each objItem In objQueryResultOperatingSystem
			If i = 1 Then
				strText = strText & objItem.Caption & ";"
				strText = strText & "'" & objItem.Version & "'" & ";"
				'strText = strText & objItem.OSArchitecture & ";"
				strText = strText & objItem.CSName & ";"
			End If

			i = i + 1
		Next

		GetCSV = strHeader & VbCrLf & strText
	End Function

End Class

'Dim objHW
'Set objHW = (New HardwareInfo)()

'WScript.Echo objHW.GetText()

'Set objHW = Nothing