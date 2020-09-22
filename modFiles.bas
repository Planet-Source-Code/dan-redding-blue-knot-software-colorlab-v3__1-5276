Attribute VB_Name = "modFiles"
Option Explicit

'Public Function getFullName(strFile As String, strExt As String)
'    getFullName = strFile & IIf(strExt = "", "", "." & strExt)
'End Function

Public Function getFullPath(strPath As String, strFile As String) As String

    If Len(strPath) = 3 Then
        getFullPath = strPath & strFile
    Else
        getFullPath = strPath & "\" & strFile
    End If
End Function
'Public Function findLastChar(strFile As String, strChar As String) As Integer
'Dim intLoop As Integer, intPos As Integer, intTest As Integer
'
'    intPos = InStr(strFile, strChar)
'    intTest = intPos
'
'    Do While intTest > 0 And intTest < Len(strFile)
'        intTest = InStr(intPos + 1, strFile, strChar)
'        If intTest <> 0 Then intPos = intTest
'    Loop
'
'    findLastChar = intPos
'
'End Function
'
'Public Function getFileExt(strFile As String) As String
'Dim intDot As Integer
'
'    intDot = findLastChar(strFile, ".")
'
'    If intDot > 0 Then
'        getFileExt = Right$(strFile, Len(strFile) - intDot)
'    Else
'        getFileExt = ""
'    End If
'
'End Function
'
'Public Function getFileFromPath(strFile As String) As String
'Dim intSlash As Integer
'
'    intSlash = findLastChar(strFile, "\")
'
'    If intSlash > 0 Then
'        getFileFromPath = Right$(strFile, Len(strFile) - intSlash)
'    Else
'        getFileFromPath = strFile
'    End If
'End Function
'
'Public Function getPathFromFullPath(strFullPath As String) As String
'Dim intSlash As Integer
'
'    intSlash = findLastChar(strFullPath, "\")
'
'    If intSlash > 0 Then
'        getPathFromFullPath = left$(strFullPath, intSlash)
'    Else
'        getPathFromFullPath = ""
'    End If
'
'End Function
