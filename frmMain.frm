VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   130
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdButton 
      Caption         =   "#"
      Height          =   375
      Index           =   3
      Left            =   3840
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "#"
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "#"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "#"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblText 
      Caption         =   "#"
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdButton_Click(Index As Integer)
    iReturn = CInt(cmdButton(Index).Tag) ' return the button's tag (the button return
    Unload Me                            ' back to the class) and close
End Sub

