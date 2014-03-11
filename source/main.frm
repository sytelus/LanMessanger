VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Messanger"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3585
   Icon            =   "main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   3585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   315
      Left            =   2460
      TabIndex        =   4
      Top             =   2100
      Width           =   975
   End
   Begin VB.TextBox txtMessage 
      Height          =   1335
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   660
      Width           =   2595
   End
   Begin VB.TextBox txtTo 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Text            =   "Person's login name"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Message:"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "To:"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   180
      Width           =   240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function NetMessageBufferSend Lib "NetApi32.dll" (ByVal sServerName As Long, ByVal sMessageName As Long, ByVal sFromName As Long, ByVal sBuff As Long, ByVal byLong As Byte) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Function SendNTMessage(ByVal vaToUserOrCompName As Variant, ByVal vsMessage As String)
    Dim lAPIRet As String
    Dim sErrorMessage As String
    Dim sToName As String
    
    Me.MousePointer = vbHourglass
    
    If VarType(vaToUserOrCompName) And vbArray <> 0 Then
    
        Dim lToIndex As Long
        
        For lToIndex = LBound(vaToUserOrCompName) To UBound(vaToUserOrCompName)
        
            sToName = Trim(vaToUserOrCompName(lToIndex))
            
            If sToName <> vbNullString Then
            
                Dim sSender As String
                sSender = GetCurUserLoginName
                If LCase$(sSender) <> "shitals" Then
                    vsMessage = vsMessage & vbCrLf & "Sent by " & sSender
                End If
            
                lAPIRet = NetMessageBufferSend(0, StrPtr(sToName), 0, StrPtr(vsMessage & " "), LenB(vsMessage))
                
                DoEvents
                
                Select Case lAPIRet
                    Case 2273
                        sErrorMessage = "Error: Message didn't reached to '" & sToName & "'." & vbCrLf & "The To field should contains valid login name or computer name or * or user might not have logged on"
                    Case Else
                        sErrorMessage = "Oops! Something wrong! To field should contain user or computer name or contact Shital! Error code is " & lAPIRet
                End Select
                
                If lAPIRet <> 0 Then
                    If MsgBox(sErrorMessage & vbCrLf & "Continue?", vbYesNo) = vbNo Then
                        Exit For
                    End If
                End If
                
            End If
            
        Next lToIndex
        
    End If
    
    Me.MousePointer = vbDefault
    
    
End Function

Private Sub cmdSend_Click()

    Dim aTo As Variant
    Dim sTo As String
    Dim lNextPos As Long
    Dim lThisPos As Long
    Dim lToLen As Long
    
    
    If LenB(txtMessage) < 256 Then
        sTo = txtTo.Text
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
            
        End If
        Call SaveSetting("Messanger", "Last Values", "To", sTo)
        Call SendNTMessage(aTo, txtMessage.Text)
    Else
        MsgBox "Message length is " & LenB(txtMessage) & " chars which exceeds permitted length of 128 char."
    End If
End Sub

Private Sub Command1_Click()
    MsgBox GetCurUserLoginName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbEnter And ((Shift And vbCtrlMask) <> 0) Then
        Call cmdSend_Click
    End If
End Sub

Private Sub Form_Load()
    txtTo.Text = GetSetting("Messanger", "Last Values", "To", "Login names seperated by semicolon (or enter * for all)")
'    txtTo.SelStart = 0
'    txtTo.SelLength = Len(txtTo.Text)
End Sub

Private Function GetCurUserLoginName() As String
    Dim sUserName As String
    Dim lSize As Long
    sUserName = Space$(255)
    lSize = Len(sUserName)
    Call GetUserName(sUserName, lSize)
    sUserName = Left(sUserName, lSize - 1)

    GetCurUserLoginName = sUserName
End Function

