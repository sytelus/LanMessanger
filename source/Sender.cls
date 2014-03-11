VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function NetMessageBufferSend Lib "NetApi32.dll" (ByVal sServerName As Long, ByVal sMessageName As Long, ByVal sFromName As Long, ByVal sBuff As Long, ByVal byLong As Byte) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function SendNTMessage(ByVal vsTo As String, ByVal vsMessage As String, Optional ByVal vsFrom As String = vbNullString) As String
    Dim lAPIRet As String
    Dim sErrorMessage As String
    Dim sToName As String
    Dim aToUserOrCompName As Variant
    
    aToUserOrCompName = DelimitedStringToArray(vsTo)
    
    If VarType(aToUserOrCompName) And vbArray <> 0 Then
    
        Dim lToIndex As Long
        
        For lToIndex = LBound(aToUserOrCompName) To UBound(aToUserOrCompName)
        
            sToName = Trim(aToUserOrCompName(lToIndex))
            
            If sToName <> vbNullString Then
            
                Dim sSender As String
                If vsFrom = vbNullString Then
                    sSender = GetCurUserLoginName
                Else
                    sSender = vsFrom
                End If
                If LCase$(sSender) <> "shitals" Then
                    vsMessage = vsMessage & vbCrLf & "Sent by " & sSender & vbCrLf & "http://msgch5532"
                    If Len(vsMessage) > 255 Then
                        vsMessage = Left(vsMessage, 252) & "..."
                    End If
                End If
            
                lAPIRet = NetMessageBufferSend(0, StrPtr(sToName), 0, StrPtr(vsMessage & " "), LenB(vsMessage))
                
                If lAPIRet <> 0 Then
                    Select Case lAPIRet
                        Case 2273
                            sErrorMessage = sErrorMessage & "Error: Message didn't reached to '" & sToName & "'." & vbCrLf & "The To field should contains valid login name or computer name or * or user might not have logged on" & vbCrLf
                        Case Else
                            sErrorMessage = sErrorMessage & "Oops! Something wrong! To field should contain user or computer name or contact Shital! Error code is " & lAPIRet & vbCrLf
                    End Select
                End If
                
            End If
            
        Next lToIndex
        
    End If
    
    SendNTMessage = sErrorMessage
    
End Function

Private Function GetCurUserLoginName() As String
    Dim sUserName As String
    Dim lSize As Long
    sUserName = Space$(255)
    lSize = Len(sUserName)
    Call GetUserName(sUserName, lSize)
    sUserName = Left(sUserName, lSize - 1)

    GetCurUserLoginName = sUserName
End Function

Private Function DelimitedStringToArray(ByVal vsTo As String) As Variant
    Dim aTo As Variant
    Dim sTo As String
    Dim lNextPos As Long
    Dim lThisPos As Long
    Dim lToLen As Long
    
    
    sTo = vsTo
    lToLen = Len(sTo)
    'If to str is not null
    If lToLen <> 0 Then
    
        ReDim aTo(1 To 1)
    
        lThisPos = 1
        lNextPos = InStr(1, sTo, ";")
        
        'If there is no delim in it
        If lNextPos = 0 Then
            aTo(1) = sTo
        Else
            
            Dim lToCount As Long
            lToCount = 0
            
            Do
                lToCount = lToCount + 1
                ReDim Preserve aTo(1 To lToCount)
                aTo(lToCount) = Mid(sTo, lThisPos, lNextPos - lThisPos)
                lThisPos = lNextPos + 1
                lNextPos = InStr(lThisPos, sTo, ";")
            Loop While lNextPos <> 0
            If lThisPos < Len(sTo) Then
                lToCount = lToCount + 1
                ReDim Preserve aTo(1 To lToCount)
                aTo(lToCount) = Mid(sTo, lThisPos)
            End If
        End If
        DelimitedStringToArray = aTo
    End If
End Function
