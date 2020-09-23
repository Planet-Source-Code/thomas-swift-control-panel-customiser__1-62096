Attribute VB_Name = "ControlPanel"
Option Explicit
Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long
Private Type GUID
     Data1 As Long
     Data2 As Long
     Data3 As Long
     Data4(8) As Byte
End Type
Public Function CreateEntryToSystemPanels(GUID As String, Titel As String, ToolTipText As String, IconDatei As String, FileToOpen As String)
Dim sKey As String
sKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\"
  CreateKey "HKEY_CLASSES_ROOT\CLSID\" & GUID
  CreateKey "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\ShellFolder"
  CreateKey "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\DefaultIcon"
  CreateKey "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\InProcServer32"
  CreateKey "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\Shell"
  CreateKey "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\Shell\Open"
  CreateKey "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\Shell\Open\Command"
  CreateKey "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\ShellFolder"
  If Form1.CheckControl.Value = 1 Then
  CreateKey "HKEY_LOCAL_MACHINE\" & sKey & "\ControlPanel\NameSpace\" & GUID
  End If
  If Form1.CheckAddDesk.Value = 1 Then
     CreateKey "HKEY_LOCAL_MACHINE\" & sKey & "\Desktop\NameSpace\" & GUID
  End If
  If Form1.CheckMyComputer.Value = 1 Then
     CreateKey "HKEY_LOCAL_MACHINE\" & sKey & "\MyComputer\NameSpace\" & GUID
  End If
    
  SetStringValue "HKEY_CLASSES_ROOT\CLSID\" & GUID, "", Titel
  SetStringValue "HKEY_CLASSES_ROOT\CLSID\" & GUID, "InfoTip", ToolTipText
  SetStringValue "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\DefaultIcon", "", IconDatei
  SetStringValue "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\InProcServer32", "", "shell32.dll"
  SetStringValue "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\InProcServer32", "ThreadingModel", "Apartment"
  SetStringValue "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\Shell\Open\Command", "", FileToOpen
  SetBinaryValue "HKEY_CLASSES_ROOT\CLSID\" & GUID & "\ShellFolder", "Attributes", Chr$(&H0) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
  If Form1.CheckControl.Value = 1 Then
  SetStringValue "HKEY_LOCAL_MACHINE" & sKey & "ControlPanel\NameSpace\" & GUID, "", ""
  End If
  If Form1.CheckAddDesk.Value = 1 Then
  SetStringValue "HKEY_LOCAL_MACHINE" & sKey & "Desktop\NameSpace\" & GUID, "", ""
  End If
  If Form1.CheckMyComputer.Value = 1 Then
  SetStringValue "HKEY_LOCAL_MACHINE" & sKey & "MyComputer\NameSpace\" & GUID, "", ""
  End If
End Function
Public Function DeleteEntryFromSystemPanels(GUID As String)
Dim sKey As String
sKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\"
  DeleteKey "HKEY_CLASSES_ROOT\CLSID\" & GUID
  DeleteKey "HKEY_LOCAL_MACHINE\" & sKey & "\Desktop\NameSpace\" & GUID
  DeleteKey "HKEY_LOCAL_MACHINE\" & sKey & "\ControlPanel\NameSpace\" & GUID
  DeleteKey "HKEY_LOCAL_MACHINE\" & sKey & "\MyComputer\NameSpace\" & GUID
End Function
Public Function GenerateUUID() As String
Dim udtGUID As GUID
Dim strGUID As String
Dim bytGUID() As Byte
Dim lngLen As Long
Dim lngRetVal As Long
Dim lngPos As Long
lngLen = 40
bytGUID = String(lngLen, 0)
CoCreateGuid udtGUID
lngRetVal = StringFromGUID2(udtGUID, VarPtr(bytGUID(0)), lngLen)
strGUID = bytGUID
If (Asc(Mid$(strGUID, lngRetVal, 1)) = 0) Then
    lngRetVal = lngRetVal - 1
End If
strGUID = Left$(strGUID, lngRetVal)
GenerateUUID = strGUID
End Function


'Copy the next code under a button
'CreateEntryToSystemPanels "{9d6D8ED6-116D-4D4E-B1C2-87098DB509BA}", "Application Name", "Tool Tipp", App.Path & "\" & "Yourapplication.exe,0", App.Path & "\" & "Yourapplication.exe -options"

'to delete the entry
'DeleteEntryFromSystemPanels "{9d6D8ED6-116D-4D4E-B1C2-87098DB509BA}"
'Have fun :-)

