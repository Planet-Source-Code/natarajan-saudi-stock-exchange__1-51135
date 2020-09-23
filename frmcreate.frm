VERSION 5.00
Begin VB.Form frmcreate 
   BackColor       =   &H0080C0FF&
   Caption         =   "Create Portfolio"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3225
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtopenbal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtname 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Open balance"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Portfolio Name"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   2415
   End
End
Attribute VB_Name = "frmcreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
Unload frmcreate
End Sub

Private Sub cmdsave_Click()
Dim portname As String, curtm As String, portmaindb As ADODB.Recordset
Dim portno As Double
portname = UCase(Trim(txtname.Text))
If Len(Trim(portname)) > 0 Then
    Set portmaindb = New ADODB.Recordset
    portmaindb.Open "select * from portfolios where portfolioname = '" & Trim(UCase(portname)) & "' ", scnn, adOpenKeyset, adLockOptimistic
    If portmaindb.RecordCount > 0 Then
        MsgBox portname & " already exisits, choose another name"
        txtname.Text = ""
        txtname.SetFocus
        Exit Sub
    End If
    portmaindb.Close
    scnn.Execute "Create table " & portname & " (compcode Number,compname text(200), stock number,avgcost number,cost number,mktprice number,mktval number,gain number,gainper number)"
    Set portmaindb = New ADODB.Recordset
    portmaindb.Open "select * from portfolios order by portfoliono", scnn, adOpenKeyset, adLockOptimistic
    If portmaindb.RecordCount > 0 Then
        portmaindb.MoveLast
        portno = portmaindb("portfoliono") + 1
    Else
        portno = 1
    End If
    portmaindb.AddNew
    portmaindb("portfoliono") = portno
    portmaindb("portfolioname") = portname
    portmaindb("cashopen") = CDbl(txtopenbal.Text)
    portmaindb("cashtran") = 0
    portmaindb("cashbal") = CDbl(txtopenbal.Text)
    portmaindb("cost") = 0
    portmaindb("mktval") = 0
    portmaindb("gainper") = 0
    portmaindb.Update
    portmaindb.Close
    MsgBox "portfolio " & portname & " created ..go to portfolio view and add transactions"
    Unload frmcreate
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
Private Sub txtopenbal_LostFocus()
If Len(txtopenbal.Text) > 0 Then
    If IsNumeric(txtopenbal.Text) Then
        cmdsave.Enabled = True
    Else
        cmdsave.Enabled = False
    End If
Else
    cmdsave.Enabled = False
End If
End Sub
