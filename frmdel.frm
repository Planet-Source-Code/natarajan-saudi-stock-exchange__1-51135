VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmdel 
   BackColor       =   &H0080C0FF&
   Caption         =   "Delete Portfolio !!!"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Portfolio"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin MSForms.ComboBox cmbSport 
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   4095
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "7223;873"
      ColumnCount     =   2
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Simplified Arabic"
      FontHeight      =   225
      FontCharSet     =   178
      FontPitchAndFamily=   2
      Object.Width           =   "5291;1058"
   End
End
Attribute VB_Name = "frmdel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload frmdel
End Sub

Private Sub cmddel_Click()
If cmbSport.ListIndex > 0 Then
    If MsgBox("wish to delete portfolio " & cmbSport.Text, vbCritical + vbYesNo, "Delete Portfolio") = vbYes Then
        scnn.Execute "drop table " & cmbSport.Text
        scnn.Execute "delete from portfolios where portfolioname = '" & cmbSport.Text & "'"
        MsgBox "portfolio deleted"
        Call cmdclose_Click
    End If
End If
End Sub
Private Sub Form_Load()
Call fillcomb(Me.cmbSport, 0, 1)
End Sub
