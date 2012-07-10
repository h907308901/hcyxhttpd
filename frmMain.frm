VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "hcyxHTTPd by h907308901"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8550
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   8550
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtMaxCon 
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Text            =   "100"
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtPort 
      Height          =   390
      Left            =   1080
      TabIndex        =   3
      Text            =   "80"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label3 
      Caption         =   "MaxCon:"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "RootDir:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    txtPath.Enabled = True
    txtPort.Enabled = True
    txtMaxCon.Enabled = True
    cmdOpen.Enabled = True
    cmdClose.Enabled = False
    HttpClose
End Sub

Private Sub cmdOpen_Click()
    txtPath.Enabled = False
    txtPort.Enabled = False
    txtMaxCon.Enabled = False
    cmdOpen.Enabled = False
    cmdClose.Enabled = True
    HttpOpen txtPath, txtMaxCon, txtPort
End Sub

Private Sub Form_Load()
    cmdClose.Enabled = False
    If Right$(App.Path, 1) = "\" Then txtPath = App.Path & "web" Else txtPath = App.Path & "\web"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdClose.Enabled Then
        If MsgBox("Are you sure to exit while running?", vbYesNo Or vbDefaultButton2) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    End
End Sub
