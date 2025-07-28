VERSION 5.00
Begin VB.Form frmMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "delaycmd"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMsg.frx":000C
   ScaleHeight     =   1770
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrExit 
      Interval        =   7000
      Left            =   180
      Top             =   1260
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   360
      Left            =   4320
      TabIndex        =   1
      Top             =   1290
      Width           =   1155
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Msg"
      Height          =   675
      Left            =   1080
      TabIndex        =   0
      Top             =   420
      Width           =   4395
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub tmrExit_Timer()
    Unload Me
End Sub
