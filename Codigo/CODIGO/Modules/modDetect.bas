Attribute VB_Name = "modDetect"
Private Const MAX_PATHLEN = &H104
Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
exeFilename(1 To MAX_PATHLEN) As Byte
End Type

Private Declare Function CloseHandle Lib "kernel32" (hObject As Long) As Boolean
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" _
(ByVal dwFlags As Long, ByVal dprocess As Long) As Long
Private Declare Function Process32First Lib "kernel32" _
(ByVal hSnapshot As Long, pProcessEntry As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" _
(ByVal hSnapshot As Long, pProcessEntry As PROCESSENTRY32) As Long
Private Const TH32CS_SNAPPROCESS = &H2&

Public Function Detected(ByVal abuscar As String) As Boolean
Dim SnapHandle, f As Long
Dim Process As PROCESSENTRY32

Detected = False

SnapHandle = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)

If SnapHandle = -1 Then Exit Function

Process.dwSize = Len(Process)
f = Process32First(SnapHandle, Process)

Do While f
If StrComp(abuscar, strZtostr(Process.exeFilename), vbTextCompare) = 0 Then
Detected = True
Exit Do
End If

Process.dwSize = Len(Process)
f = Process32Next(SnapHandle, Process)
Loop

Call CloseHandle(SnapHandle)

End Function

Private Function strZtostr(Fuente() As Byte) As String

Dim i&, cb As Byte
strZtostr = ""
Do While i < MAX_PATHLEN
i = i + 1
cb = Fuente(i)
If cb = 0 Then Exit Function 'Return
strZtostr = strZtostr & Chr$(cb)
Loop

End Function

