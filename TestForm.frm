VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "cMsgBox Testing Pad"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TestForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   401
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Text            =   "Title"
      Top             =   5640
      Width           =   3135
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   12
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Frame fraIcon 
      Appearance      =   0  'Flat
      Caption         =   "Icon"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   3360
      TabIndex        =   8
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select Icon"
         Height          =   615
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear Icon"
         Height          =   615
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblPreview 
         Caption         =   "Icon [ No Icon ]"
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   360
         Width           =   2535
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   480
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
   End
   Begin MSComDlg.CommonDialog cmdIcon 
      Left            =   3360
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select Icon"
      Filter          =   "Icon Files (ico cur) | *.ico;*.cur"
   End
   Begin VB.TextBox txtPromptText 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "TestForm.frx":0442
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Frame fraCustom 
      Appearance      =   0  'Flat
      Caption         =   "Custom Button Text"
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   3135
      Begin VB.TextBox txtCustom 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Text            =   "Custom 4"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtCustom 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Text            =   "Custom 3"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtCustom 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Text            =   "Custom 2"
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtCustom 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Text            =   "Custom 1"
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame fraStyles 
      Appearance      =   0  'Flat
      Caption         =   "Button Styles"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstStyles 
         Appearance      =   0  'Flat
         Height          =   1335
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Label lblReturn 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   4800
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cMsg As cMsgBox
Dim sReturn() As Variant

Private Sub cmdClear_Click()
        imgIcon.Picture = LoadPicture("")
        lblPreview.Caption = "Icon [ No Icon ]"
End Sub

Private Sub cmdPreview_Click()
    Dim iReturn As Integer
    cMsg.CustomText1 = txtCustom(0).Text
    cMsg.CustomText2 = txtCustom(1).Text
    cMsg.CustomText3 = txtCustom(2).Text
    cMsg.CustomText4 = txtCustom(3).Text
    iReturn = cMsg.cMsgBox(txtPromptText.Text, lstStyles.ListIndex + 1, txtTitle.Text, imgIcon.Picture)
    
    lblReturn.Caption = "The button clicked was: " & sReturn(iReturn) & " (Constant: c" & sReturn(iReturn) & ")"
End Sub

Private Sub cmdSelect_Click()
    cmdIcon.ShowOpen
    If cmdIcon.FileName <> "" Then
        imgIcon.Picture = LoadPicture(cmdIcon.FileName)
        lblPreview.Caption = "Icon [ " & cmdIcon.FileTitle & " ]"
    End If
End Sub

Private Sub Form_Load()
    Set cMsg = New cMsgBox
    sReturn = Array("", "OK", "Cancel", "Yes", "No", "Abort", "Retry", "Ignore", _
        "Custom Button 1", "Custom Button 2", "Custom Button 3", "Custom Button 4")
    With lstStyles
        .AddItem "OK Only"
        .AddItem "OK and Cancel"
        .AddItem "Yes and No"
        .AddItem "Yes, No and Cancel"
        .AddItem "Abort, Retry and Ignore"
        .AddItem "Custom Buttons"
    End With
    lstStyles.Selected(0) = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cMsg = Nothing
End Sub

Private Sub lblTitles_Click(Index As Integer)

End Sub

