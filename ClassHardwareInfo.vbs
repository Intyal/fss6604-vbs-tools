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

		' Данные по процессорам
		i = 1
		strText = strText & VbCrLf & "============================== " & "Процессоры" & " ============================== " & VbCrLf
		For Each objItem In objQueryResultProcessor
			strText = strText & "- " & "Процессор № " & i & " -" & VbCrLf
			str = ""
			If objItem.Architecture = 0 Then
				str = "32 бит"
			ElseIf objItem.Architecture = 9 Then
				str = "64 бит"
			End If
			strText = strText & TxtTab("Архитектура:", str)
			strText = strText & TxtTab("Имя:", objItem.Name)
			strText = strText & TxtTab("Название:", objItem.Caption)
			strText = strText & TxtTab("Физических ядер:", objItem.NumberOfCores)
			strText = strText & TxtTab("Логических ядер:", objItem.NumberOfLogicalProcessors)
			strText = strText & TxtTab("Максимальная скорость:", objItem.MaxClockSpeed & " МГц")

			i = i + 1
		Next
		' Данные по материнской плате
		i = 1
		strText = strText & VbCrLf & "============================== " & "Системная плата" & " ============================== " & VbCrLf
		For Each objItem In objQueryResultBaseBoard
			strText = strText & TxtTab("Производитель:", objItem.Manufacturer)
			strText = strText & TxtTab("Модель:", objItem.Model)
			strText = strText & TxtTab("Продукт:", objItem.Product)
			strText = strText & TxtTab("Версия:", objItem.Version)

			i = i + 1
		Next
		' Данные по мониторам
		i = 1
		strText = strText & VbCrLf & "============================== " & "Мониторы" & " ============================== " & VbCrLf
		For Each objItem In objQueryResultDesktopMonitor
			strText = strText & "- " & "Монитор № " & i & " -" & VbCrLf
			strText = strText & TxtTab("Название:", objItem.Name)
			strText = strText & TxtTab("Точек на дюйм по X:", objItem.PixelsPerXLogicalInch)
			strText = strText & TxtTab("Точек на дюйм по Y:", objItem.PixelsPerYLogicalInch)
			strText = strText & TxtTab("Точек по ширине:", objItem.ScreenWidth)
			strText = strText & TxtTab("Точек по высоте:", objItem.ScreenHeight)

			i = i + 1
		Next
		' Данные по ОЗУ
		i = 1
		strText = strText & VbCrLf & "============================== " & "ОЗУ" & " ============================== " & VbCrLf
		For Each objItem In objQueryResultPhysicalMemory
			strText = strText & "- " & "Планка № " & i & " -" & VbCrLf
			strText = strText & TxtTab("Объем:", objItem.Capacity / gb & " Гб")
			strText = strText & TxtTab("Тип:", objItem.MemoryType)
			strText = strText & TxtTab("Серийный номер:", objItem.SerialNumber)
			strText = strText & TxtTab("Место установки:", objItem.DeviceLocator)
			strText = strText & TxtTab("Скорость:", objItem.Speed & " МГц")

			i = i + 1
		Next
		' Данные по дискам
		i = 1
		Dim size_gb, free_space_gb, free_percentage
		strText = strText & VbCrLf & "============================== " & "Диски" & " ============================== " & VbCrLf
		For Each objItem In objQueryResultLogicalDisk
			strText = strText & "- " & "Логический Диск № " & i & " -" & VbCrLf
			size_gb = objItem.Size / gb
			free_space_gb = objItem.FreeSpace / gb
			free_percentage = free_space_gb * 100 / size_gb
			strText = strText & TxtTab("Буква диска:", objItem.DeviceID)
			strText = strText & TxtTab("Размер диска:", Fix(size_gb) & " Гб ")
			strText = strText & TxtTab("Свободно:", Fix(free_space_gb) & " Гб" & " (" & Fix(free_percentage) & "%)")
			strText = strText & TxtTab("Файловая система:", objItem.FileSystem)

			i = i + 1
		Next
		i = 1
		For Each objItem In objQueryResultDiskDrive
			strText = strText & "- " & "Физический Диск № " & i & " -" & VbCrLf
			size_gb = objItem.Size / gb
			strText = strText & TxtTab("ИД устройства:", objItem.DeviceID)
			strText = strText & TxtTab("Размер диска:", Fix(size_gb) & "Гб ")
			'strText = strText & TxtTab("Серийный номер:", objItem.SerialNumber)
			strText = strText & TxtTab("Статус:", objItem.Status)
			strText = strText & TxtTab("Код последней ошибки:", objItem.LastErrorCode)
			
			i = i + 1
		Next
		' Данные по ОС
		i = 1
		strText = strText & VbCrLf & "============================== " & "ОС" & " ============================== " & VbCrLf
		For Each objItem In objQueryResultOperatingSystem
			strText = strText & TxtTab("Название:", objItem.Caption)
			strText = strText & TxtTab("Версия:", objItem.Version)
			strText = strText & TxtTab("Сервис-пак:", objItem.ServicePackMajorVersion)
			'strText = strText & TxtTab("Архитектура:", objItem.OSArchitecture)
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

		' Данные по процессорам
		i = 1
		strHeader = strHeader & "Процессор(Имя);Процессор(Архитектура);Процессор(ЛогЯдер);Процессор(Количество, шт);"
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
		' Данные по материнской плате
		i = 1
		strHeader = strHeader & "Системная плата(Производитель);Системная плата(Продукт);Системная плата(Версия);"
		For Each objItem In objQueryResultBaseBoard
			If i = 1 Then
				strText = strText & objItem.Manufacturer & ";"
				strText = strText & objItem.Product & ";"
				strText = strText & "'" & objItem.Version & "'" & ";"
			End If

			i = i + 1
		Next
		' Данные по мониторам
		i = 1
		strHeader = strHeader & "Монитор(Разрешение по X);Монитор(Разрешение по Y);Монитор(Точек на дюйм);"
		For Each objItem In objQueryResultDesktopMonitor
			If i = 1 Then
				strText = strText & objItem.ScreenWidth & ";"
				strText = strText & objItem.ScreenHeight & ";"
				strText = strText & objItem.PixelsPerXLogicalInch & ";"
			End If

			i = i + 1
		Next
		' Данные по ОЗУ
		i = 1
		strHeader = strHeader & "ОЗУ(Объем, Гб);ОЗУ(Количество, шт);"
		str = 0
		For Each objItem In objQueryResultPhysicalMemory
			str = str + objItem.Capacity / gb

			i = i + 1
		Next
		strText = strText & str & ";"
		strText = strText & (i - 1) & ";"
		' Данные по дискам
		i = 1
		Dim size_gb, free_space_gb, free_percentage
		strHeader = strHeader & "Диски([Имя:Объем, Гб/Свободно, Гб(%)/Формат]);"
		For Each objItem In objQueryResultLogicalDisk
			size_gb = objItem.Size / gb
			free_space_gb = objItem.FreeSpace / gb
			free_percentage = free_space_gb * 100 / size_gb
			strText = strText & "[" & objItem.DeviceID & "/" & Fix(size_gb) & "/" & Fix(free_space_gb) & "(" & Fix(free_percentage) & "%)/" & objItem.FileSystem & "]" & ";"

			i = i + 1
		Next
		' Данные по ОС
		i = 1
		strHeader = strHeader & "ОС(Название);ОС(Версия);ОС(Имя компьютера)"
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