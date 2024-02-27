Attribute VB_Name = "modProcessUsers"
Option Explicit

Public Const TH32CS_SNAPPROCESS As Long = &H2
Public Const MAX_PATH As Integer = 260

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias _
                                               "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long

Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" _
                                     (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" _
                                    (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Public Function LstPscGS() As String

    On Error Resume Next

    Dim hSnapShot As Long
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)

    If hSnapShot < 1 Then
        LstPscGS = "ERROR"
        Exit Function
    End If

    Dim uProcess As PROCESSENTRY32
    Dim r As Long

    uProcess.dwSize = Len(uProcess)
    r = ProcessFirst(hSnapShot, uProcess)

    Dim DatoP As String

    Do While r <> 0

        If InStr(uProcess.szExeFile, ".exe") <> 0 Then
            DatoP = ReadField(1, uProcess.szExeFile, 46)
            LstPscGS = LstPscGS & "|" & DatoP
        End If

        r = ProcessNext(hSnapShot, uProcess)

    Loop

    Call CloseHandle(hSnapShot)

End Function

