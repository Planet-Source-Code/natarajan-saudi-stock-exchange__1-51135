VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmcurrentport 
   BackColor       =   &H00C0FFC0&
   ClientHeight    =   6090
   ClientLeft      =   -180
   ClientTop       =   150
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid msgrd 
      Height          =   5655
      Left            =   0
      TabIndex        =   17
      Top             =   1800
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9975
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      BackColor       =   -2147483624
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483628
      BackColorBkg    =   8438015
      TextStyle       =   4
      TextStyleFixed  =   3
      FocusRect       =   2
      HighLight       =   2
      GridLines       =   2
      GridLinesFixed  =   3
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Simplified Arabic"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.Frame Frame2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Portfolio details .....press F5 to refresh mkt price at anytime"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   5640
         TabIndex        =   4
         Top             =   120
         Width           =   6135
         Begin VB.TextBox txtgain 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   4440
            TabIndex        =   16
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txtmkt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   4440
            TabIndex        =   14
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtcost 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   4440
            TabIndex        =   12
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtbal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   1200
            TabIndex        =   10
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txttran 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   1200
            TabIndex        =   8
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtopen 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   1200
            TabIndex        =   6
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Gain/Loss"
            BeginProperty Font 
               Name            =   "Simplified Arabic"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   3360
            TabIndex        =   15
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Mkt Value"
            BeginProperty Font 
               Name            =   "Simplified Arabic"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   3360
            TabIndex        =   13
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Cost basis"
            BeginProperty Font 
               Name            =   "Simplified Arabic"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   3360
            TabIndex        =   11
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Cash Bal."
            BeginProperty Font 
               Name            =   "Simplified Arabic"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Cash Tran"
            BeginProperty Font 
               Name            =   "Simplified Arabic"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cash Open"
            BeginProperty Font 
               Name            =   "Simplified Arabic"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   65535
         Left            =   120
         Top             =   1680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Portfolio"
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
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin MSForms.ComboBox cmbport 
         Height          =   495
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   4215
         VariousPropertyBits=   746604571
         BackColor       =   8438015
         DisplayStyle    =   3
         Size            =   "7435;873"
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
   Begin VB.Label lbltemp 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuaddtran 
         Caption         =   "Add Transaction"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuupdprice 
         Caption         =   "Update prices"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuclose 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuoption 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu mnuprn 
         Caption         =   "Print"
      End
      Begin VB.Menu mnutran 
         Caption         =   "Add Tran"
      End
      Begin VB.Menu mnurefresh 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "frmcurrentport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim stockdb As ADODB.Recordset
Dim portno As Double
Private Sub cmbport_Click()
Call updcomp
End Sub
Private Sub Form_Load()
Dim sql As String
Call fillcomb(Me.cmbport, 0, 1)
Call updcomp
End Sub
Sub updcomp()
Dim sql As String
If cmbport.ListIndex > 0 Then
    sql = "select * from " & cmbport.Text & " order by compcode "
    Set stockdb = New ADODB.Recordset
    stockdb.Open sql, scnn, adOpenKeyset, adLockReadOnly
    Call updportdtls
    Call grd
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
stockdb.Close
End Sub
Private Sub mnuaddtran_Click()
frmtran.Show
End Sub
Private Sub mnuclose_Click()
On Error Resume Next
stockdb.Close
Unload frmcurrentport
End Sub
Private Sub mnuprint_Click()
If stockdb.RecordCount > 0 Then
    With dsrpandl1
        Set .DataSource = stockdb
        .Sections("section4").Controls("label2").Caption = cmbport.Text
        .Show vbModal
    End With
Else
    MsgBox "no records", , "Portfolio print"
End If
End Sub
Private Sub mnuprn_Click()
Call mnuprint_Click
End Sub
Private Sub mnurefresh_Click()
Call mnuupdprice_Click
End Sub
Private Sub mnutran_Click()
frmtran.Show
End Sub
Private Sub mnuupdprice_Click()
DoEvents
MDIForm1.updprices
Call updcomp
End Sub
Private Sub msgrd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    If msgrd.Rows > 0 Then
        msgrd.Sort = msgrd.Col
    End If
ElseIf Button = 2 Then
    PopupMenu mnuoption
End If
End Sub
Private Sub Timer1_Timer()
If cmbport.ListIndex > 0 Then
    stockdb.Requery
    Call updcomp
End If
End Sub
Sub updportdtls()
Dim Pdb As ADODB.Recordset
txtopen.Text = ""
txttran.Text = ""
txtbal.Text = ""
txtcost.Text = ""
txtmkt.Text = ""
txtgain.Text = ""

If cmbport.ListIndex > 0 Then
    Set Pdb = New ADODB.Recordset
    Pdb.Open "select * from portfolios where portfoliono = " & cmbport.List(cmbport.ListIndex, 1), scnn, adOpenKeyset, adLockReadOnly
    If Pdb.RecordCount > 0 Then
        txtopen.Text = IIf(IsNull(Pdb("cashopen")), 0, Pdb("cashopen"))
        txttran.Text = IIf(IsNull(Pdb("cashtran")), 0, Pdb("cashtran"))
        txtbal.Text = IIf(IsNull(Pdb("cashbal")), 0, Pdb("cashbal"))
        txtcost.Text = IIf(IsNull(Pdb("cost")), 0, Pdb("cost"))
        txtmkt.Text = IIf(IsNull(Pdb("mktval")), 0, Pdb("mktval"))
        txtgain.Text = CDbl(txtmkt.Text) - CDbl(txtcost.Text)
    End If
    Pdb.Close
End If
End Sub
Sub grd()
Dim i, j
With msgrd
   .Col = 0
   .Row = 0
   .Text = "Code"
   .ColWidth(0) = 735
   .Col = 1
   .Text = "Company"
   .ColWidth(1) = 1700
   .Col = 2
   .Text = "Holding"
   .ColWidth(2) = 1390
   .Col = 3
   .Text = "Avg. Cost"
   .ColWidth(3) = 1065
   .Col = 4
   .Text = "Cost Basis"
   .ColWidth(4) = 1600
   .Col = 5
   .Text = "Mkt. Price"
   .ColWidth(5) = 1065
   .Col = 6
   .Text = "Mkt. Value"
   .ColWidth(6) = 1600
   .Col = 7
   .Text = "Gain/loss"
   .ColWidth(7) = 1600
   .Col = 8
   .Text = " % "
   .ColWidth(8) = 1065
   If stockdb.RecordCount > 0 Then
        .Rows = stockdb.RecordCount + 1
        For i = 0 To (stockdb.RecordCount - 1)
            .Row = i + 1
            For j = 0 To (.Cols - 1)
                .Col = j
                If j >= 2 And j < 8 Then
                    .Text = Format(stockdb.Fields(j), "#,##0.00")
                ElseIf j = 8 Then
                    .Text = Format(stockdb.Fields(j) * 100, "##0.00")
                Else
                    .Text = Trim(stockdb.Fields(j))
                End If
            Next j
            stockdb.MoveNext
            If stockdb.EOF Then Exit For
        Next
    End If
End With
End Sub
