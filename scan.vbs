Set CommonDialog = CreateObject("WIA.CommonDialog")
Set DeviceManager = CreateObject("WIA.DeviceManager")

' List all Available Devices by Name and DeviceID
' The following example shows how to list all available Devices by name and DeviceID.
Dim i, count, id, deviceName, deviceID
count = DeviceManager.DeviceInfos.Count
WScript.Echo "Number of Devices found = " & count
For i = 1 to count
	deviceName = DeviceManager.DeviceInfos(i).Properties("Name").Value
	deviceID = DeviceManager.DeviceInfos(i).DeviceID
	WScript.Echo "Device " & i & ":" & vbTab & deviceName & vbTab & "(" & deviceID & ")"
	If InStr(deviceName, "CanoScan") Then
		id = i
	End If
Next

Set DevInfo = DeviceManager.DeviceInfos(id)
Set Device = DevInfo.Connect

WScript.Echo "Current device ID: " & id

Device.Items(1).Properties("6146").Value = WScript.Arguments.Item(1) '4 is black-white, gray is 2, color 1 (Color Intent)
Device.Items(1).Properties("6147").Value = 600 'dots per inch/horizontal
Device.Items(1).Properties("6148").Value = 600 'dots per inch/vertical
Device.Items(1).Properties("6149").Value = 0 'x point where to start scan
Device.Items(1).Properties("6150").Value = 0 'y point where to start scan
Device.Items(1).Properties("6151").Value = 5100 'horizontal exent DPI x inches wide
Device.Items(1).Properties("6152").Value = 7002 'vertical extent DPI x inches tall

'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-consts-formatid
Set img = CommonDialog.ShowTransfer(Device.Items(1), "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}", true)

img.SaveFile WScript.Arguments.Item(0)