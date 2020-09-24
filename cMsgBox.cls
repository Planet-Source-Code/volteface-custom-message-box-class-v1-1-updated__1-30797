VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************
'* cMsgBox Class
'* By VolteFace
'* (c) 2001
'*******************************************
Option Explicit
Private m_ButtonStyle As Integer
Private m_Custom1 As String
Private m_Custom2 As String
Private m_Custom3 As String
Private m_Custom4 As String
Public Enum cMsgBoxConstants
    [OK Only] = 1
    [OK and Cancel]
    [Yes and No]
    [Yes, No and Cancel]
    [Abort, Retry and Ignore]
    [Custom Buttons]
End Enum
Public Enum cMsgBoxReturns
    cOK = 1
    cCancel
    cYes
    cNo
    cAbort
    cRetry
    cIgnore
    cCustom1
    cCustom2
    cCustom3
    cCustom4
End Enum
Public Function cMsgBox(Prompt As String, Optional Style As cMsgBoxConstants = [OK Only], Optional Title As String = "Message", Optional Icon As StdPicture) As cMsgBoxReturns
    Dim i As Integer
    frmMain.lblText.Caption = Prompt
    frmMain.imgIcon.Picture = Icon
    If Not frmMain.imgIcon.Picture.Height = 0 Then
        frmMain.imgIcon.Visible = True
        frmMain.lblText.Left = frmMain.imgIcon.Left + frmMain.imgIcon.Width + 8 ' There is an icon, make the label move over
        frmMain.lblText.Width = frmMain.ScaleWidth - frmMain.imgIcon.Width - 24
    Else
        frmMain.imgIcon.Visible = False
        frmMain.lblText.Left = frmMain.imgIcon.Left + 8 ' no icon, the label can take up the entire space
        frmMain.lblText.Width = frmMain.ScaleWidth - frmMain.imgIcon.Width
    End If
    frmMain.Caption = Title ' set the title
    frmMain.lblText.Height = GetTextHeight(frmMain.lblText.Caption) ' resize the label according to the height of the text
    frmMain.Height = (frmMain.lblText.Height + frmMain.lblText.Top + _
        frmMain.cmdButton(0).Height * 2 + (frmMain.cmdButton(0).Left) * 2) * _
            Screen.TwipsPerPixelY ' resize the form to match the label
    For i = 0 To 3
        frmMain.cmdButton(i).Top = (frmMain.ScaleHeight - _
            frmMain.cmdButton(i).Height) - frmMain.cmdButton(0).Left ' reposition the buttons to the bottom of the form
    Next i
    Select Case Style ' look for different styles and change the button properties accordingly
        Case [OK Only]
            frmMain.cmdButton(3).Tag = cOK
            frmMain.cmdButton(3).Caption = "OK"
            frmMain.cmdButton(3).Visible = True
            frmMain.cmdButton(2).Visible = False
            frmMain.cmdButton(1).Visible = False
            frmMain.cmdButton(0).Visible = False
        Case [OK and Cancel]
            frmMain.cmdButton(2).Tag = cOK
            frmMain.cmdButton(2).Caption = "OK"
            frmMain.cmdButton(2).Visible = True
            frmMain.cmdButton(3).Tag = cCancel
            frmMain.cmdButton(3).Caption = "Cancel"
            frmMain.cmdButton(3).Visible = True
            frmMain.cmdButton(1).Visible = False
            frmMain.cmdButton(0).Visible = False
        Case [Yes and No]
            frmMain.cmdButton(2).Tag = cYes
            frmMain.cmdButton(2).Caption = "Yes"
            frmMain.cmdButton(2).Visible = True
            frmMain.cmdButton(3).Tag = cNo
            frmMain.cmdButton(3).Caption = "No"
            frmMain.cmdButton(3).Visible = True
            frmMain.cmdButton(1).Visible = False
            frmMain.cmdButton(0).Visible = False
        Case [Yes, No and Cancel]
            frmMain.cmdButton(3).Tag = cCancel
            frmMain.cmdButton(3).Caption = "Cancel"
            frmMain.cmdButton(3).Visible = True
            frmMain.cmdButton(1).Tag = cYes
            frmMain.cmdButton(1).Caption = "Yes"
            frmMain.cmdButton(1).Visible = True
            frmMain.cmdButton(2).Tag = cNo
            frmMain.cmdButton(2).Caption = "No"
            frmMain.cmdButton(2).Visible = True
            frmMain.cmdButton(0).Visible = False
        Case [Abort, Retry and Ignore]
            frmMain.cmdButton(1).Tag = cAbort
            frmMain.cmdButton(1).Caption = "About"
            frmMain.cmdButton(1).Visible = True
            frmMain.cmdButton(2).Tag = cRetry
            frmMain.cmdButton(2).Caption = "Retry"
            frmMain.cmdButton(2).Visible = True
            frmMain.cmdButton(3).Tag = cIgnore
            frmMain.cmdButton(3).Caption = "Ignore"
            frmMain.cmdButton(3).Visible = True
            frmMain.cmdButton(0).Visible = False
        Case [Custom Buttons]
            frmMain.cmdButton(0).Tag = cCustom1
            frmMain.cmdButton(0).Caption = m_Custom1
            frmMain.cmdButton(0).Visible = (m_Custom1 <> "")
            frmMain.cmdButton(1).Tag = cCustom2
            frmMain.cmdButton(1).Caption = m_Custom2
            frmMain.cmdButton(1).Visible = (m_Custom2 <> "")
            frmMain.cmdButton(2).Tag = cCustom3
            frmMain.cmdButton(2).Caption = m_Custom3
            frmMain.cmdButton(2).Visible = (m_Custom3 <> "")
            frmMain.cmdButton(3).Tag = cCustom4
            frmMain.cmdButton(3).Caption = m_Custom4
            frmMain.cmdButton(3).Visible = (m_Custom4 <> "")
    End Select
    frmMain.Show vbModal
    cMsgBox = iReturn ' make the cMsgBox return something
End Function
Private Function GetTextHeight(sText As String) As Integer
    '********************************************************
    '* GetTextHeight
    '* By VolteFace
    '* (c) 2001
    '********************************************************
    '* Used to get the height of word-wrapped text
    '********************************************************
    Dim sDummy As String, sTempText As String
    Dim i As Integer
    
    For i = 1 To Len(sText)
        sDummy = sDummy & Mid(sText, i, 1) ' add characters to the two temp. strings
        sTempText = sTempText & Mid(sText, i, 1)
        If frmMain.TextWidth(sDummy) >= frmMain.lblText.Width Then ' if the text needs to be wrapped...
            sTempText = sTempText & vbCrLf  ' ... add a new line
            sDummy = "" ' reset the dummy
        End If
    Next i
    
    GetTextHeight = frmMain.TextHeight(sTempText & vbCrLf) + (frmMain.TextHeight("B") * 2) ' add some padding for good measure
End Function
Public Property Get CustomText1() As String
    CustomText1 = m_Custom1
End Property
Public Property Let CustomText1(ByVal vNewValue As String)
    m_Custom1 = vNewValue
End Property
Public Property Get CustomText2() As String
    CustomText2 = m_Custom1
End Property
Public Property Let CustomText2(ByVal vNewValue As String)
    m_Custom2 = vNewValue
End Property
Public Property Get CustomText3() As String
    CustomText3 = m_Custom1
End Property
Public Property Let CustomText3(ByVal vNewValue As String)
    m_Custom3 = vNewValue
End Property
Public Property Get CustomText4() As String
    CustomText4 = m_Custom1
End Property
Public Property Let CustomText4(ByVal vNewValue As String)
    m_Custom4 = vNewValue
End Property
