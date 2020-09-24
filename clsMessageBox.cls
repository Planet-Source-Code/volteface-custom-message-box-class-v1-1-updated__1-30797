VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event ButtonClicked(Button As Integer)
Dim m_NumButtons As Integer
Dim m_ButtonText As String
Public Sub MessageBox(sText As String, Optional sTitle As String)
    Dim fMsgForm As Form
    Set fMsgForm = New Form1
    Dim sButtons() As String
    
    If m_NumButtons = 0 Then
        Err.Raise 30000, , "No buttons defined"
    End If
        sButtons = Split(m_ButtonText, ";")
    If UBound(sButtons) = 0 Then
        Err.Raise 30001, , "No button text defined"
    ElseIf UBound(sButtons) <> m_NumButtons - 1 Then
        Err.Raise 30002, , "Number of buttons does not match number of button texts"
    End If
    If IsMissing(sTitle) Then
        sTitle = App.Title
    End If
    
    fMsgForm.Show
End Sub
Public Property Get NumButtons() As Integer
    NumButtons = m_NumButtons
End Property
Public Property Let NumButtons(ByVal v_NewNumButtons As Integer)
    m_NumButtons = v_NewNumButtons
End Property
Public Property Get ButtonText() As String
    ButtonText = m_ButtonText
End Property
Public Property Let ButtonText(ByVal v_NewButtonText As String)
    m_ButtonText = v_NewButtonText
End Property

