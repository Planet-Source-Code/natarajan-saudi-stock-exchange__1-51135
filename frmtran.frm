VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmtran 
   BackColor       =   &H0080C0FF&
   Caption         =   "Transaction"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   Begin InetCtlsObjects.Inet inettran 
      Left            =   120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   10186
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Buy Stocks"
      TabPicture(0)   =   "frmtran.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtBcode"
      Tab(0).Control(1)=   "txtcoEname"
      Tab(0).Control(2)=   "txtcoAname"
      Tab(0).Control(3)=   "txtBqty"
      Tab(0).Control(4)=   "txtBcost"
      Tab(0).Control(5)=   "cmdBsave"
      Tab(0).Control(6)=   "cmdcancel(0)"
      Tab(0).Control(7)=   "Label17"
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(9)=   "Label2"
      Tab(0).Control(10)=   "Label3"
      Tab(0).Control(11)=   "Label4"
      Tab(0).Control(12)=   "Label5"
      Tab(0).Control(13)=   "Label11"
      Tab(0).Control(14)=   "cmbBport"
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Sell stocks"
      TabPicture(1)   =   "frmtran.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtsqty"
      Tab(1).Control(1)=   "txtsval"
      Tab(1).Control(2)=   "cmdSsave"
      Tab(1).Control(3)=   "cmdcancel(1)"
      Tab(1).Control(4)=   "Label6"
      Tab(1).Control(5)=   "Label9"
      Tab(1).Control(6)=   "Label10"
      Tab(1).Control(7)=   "Label12"
      Tab(1).Control(8)=   "cmbSport"
      Tab(1).Control(9)=   "cmbcomp"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Withdraw/deposit cash"
      TabPicture(2)   =   "frmtran.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label13"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmbWport"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label15"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label14"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label16"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label7"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "lsttype"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cmdcancel(2)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cmdWsave"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtWamt"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "txtcurbal"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "txtbal"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      Begin VB.TextBox txtbal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Left            =   4440
         TabIndex        =   36
         Text            =   "0"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtcurbal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Left            =   4440
         TabIndex        =   33
         Text            =   "0"
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox txtBcode 
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
         Left            =   -70560
         TabIndex        =   2
         Text            =   "1090"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtcoEname 
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
         Left            =   -70560
         TabIndex        =   3
         Top             =   1680
         Width           =   3975
      End
      Begin VB.TextBox txtcoAname 
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
         Left            =   -70560
         TabIndex        =   4
         Top             =   2160
         Width           =   3975
      End
      Begin VB.TextBox txtBqty 
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
         Left            =   -70560
         TabIndex        =   5
         Text            =   "0"
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox txtBcost 
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
         Left            =   -70560
         TabIndex        =   6
         Text            =   "0"
         Top             =   3480
         Width           =   2175
      End
      Begin VB.CommandButton cmdBsave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70440
         TabIndex        =   7
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   0
         Left            =   -69480
         TabIndex        =   8
         Top             =   5160
         Width           =   855
      End
      Begin VB.TextBox txtsqty 
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
         Left            =   -70560
         TabIndex        =   11
         Text            =   "0"
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txtsval 
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
         Left            =   -70560
         TabIndex        =   12
         Text            =   "0"
         Top             =   3600
         Width           =   1815
      End
      Begin VB.CommandButton cmdSsave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70320
         TabIndex        =   13
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   1
         Left            =   -69360
         TabIndex        =   14
         Top             =   4920
         Width           =   855
      End
      Begin VB.TextBox txtWamt 
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
         Left            =   4440
         TabIndex        =   17
         Text            =   "0"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.CommandButton cmdWsave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4680
         TabIndex        =   18
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   2
         Left            =   5760
         TabIndex        =   19
         Top             =   4920
         Width           =   855
      End
      Begin VB.ListBox lsttype 
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         ItemData        =   "frmtran.frx":0054
         Left            =   4440
         List            =   "frmtran.frx":0061
         TabIndex        =   16
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Cash Balance"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label17 
         Caption         =   "checking code ....."
         Height          =   255
         Left            =   -69240
         TabIndex        =   35
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "New Balance"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   3960
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Comp. Symbol"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   32
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Company English  Name"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   31
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Company Arabic Name"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   30
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   29
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Total Cost"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   28
         Top             =   3480
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "Select Comp. Symbol"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   27
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label9 
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   26
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Label Label10 
         Caption         =   "Sale Value"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   25
         Top             =   3600
         Width           =   3015
      End
      Begin VB.Label Label11 
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
         Left            =   -74640
         TabIndex        =   24
         Top             =   600
         Width           =   855
      End
      Begin MSForms.ComboBox cmbBport 
         Height          =   495
         Left            =   -70560
         TabIndex        =   1
         Top             =   480
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
      Begin VB.Label Label12 
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
         Left            =   -74760
         TabIndex        =   23
         Top             =   600
         Width           =   855
      End
      Begin MSForms.ComboBox cmbSport 
         Height          =   495
         Left            =   -70560
         TabIndex        =   9
         Top             =   600
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
      Begin MSForms.ComboBox cmbcomp 
         Height          =   495
         Left            =   -70560
         TabIndex        =   10
         Top             =   1200
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
      Begin VB.Label Label14 
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   2880
         Width           =   3015
      End
      Begin VB.Label Label15 
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
         Left            =   360
         TabIndex        =   21
         Top             =   600
         Width           =   855
      End
      Begin MSForms.ComboBox cmbWport 
         Height          =   495
         Left            =   4440
         TabIndex        =   15
         Top             =   600
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
      Begin VB.Label Label13 
         Caption         =   "Transaction Type"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   2280
         Width           =   3015
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "file"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmtran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Function checkcomp(tempcode As String) As Boolean
Dim thepage As String, returnstr As String
Dim exactstr As String, beg As Integer, endline As Integer
On Error GoTo errh:
DoEvents
checkcomp = False
thepage = "http://www.tadawul.com.sa/user/ListCompany.ASP"
returnstr = thepage
thepage = inettran.OpenURL(thepage, icString)
returnstr = inettran.GetChunk(2048, icString)
Do While Len(returnstr) <> 0
    DoEvents
    returnstr = inettran.GetChunk(2048, icString)
Loop

beg = 0
beg = InStr(1, thepage, tempcode)
If beg = 0 Then
    checkcomp = False
    Exit Function
End If
beg = InStr(beg, thepage, "ColLink")
If beg > 0 Then
    beg = beg + 1
    beg = InStr(beg, thepage, ">")
    endline = InStr(beg, thepage, "<")
    exactstr = Mid(thepage, beg + 1, (endline - (beg + 1)))
    txtcoEname.Text = Trim(UCase(exactstr))
    txtcoAname.Text = Trim(UCase(exactstr))
    checkcomp = True
End If
Exit Function

errh:
txtcoEname.Text = ""
txtcoAname.Text = ""
End Function
Private Sub cmbcomp_LostFocus()
Dim stockdb As ADODB.Recordset
If cmbSport.ListIndex > 0 Then
    If cmbcomp.ListIndex > 0 Then
        Set stockdb = New ADODB.Recordset
        stockdb.Open "select * from " & cmbSport.Text & _
            " where compcode = " & cmbcomp.List(cmbcomp.ListIndex, 1), scnn, adOpenKeyset, adLockReadOnly
        If stockdb.RecordCount > 0 Then
            txtsqty.Text = stockdb("stock")
            txtsval.Text = stockdb("mktval")
        Else
            MsgBox "stock not found for " & cmbcomp.Text & " in portfolio " & cmbSport.Text
            stockdb.Close
            cmbcomp.SetFocus
            Exit Sub
        End If
        stockdb.Close
    Else
        cmbcomp.SetFocus
    End If
Else
    cmbSport.SetFocus
End If
End Sub
Private Sub cmbWport_LostFocus()
Dim portdb As ADODB.Recordset
If cmbWport.ListIndex > 0 Then
    Set portdb = New ADODB.Recordset
    portdb.Open "select * from portfolios where portfoliono = " & cmbWport.List(cmbWport.ListIndex, 1), scnn, adOpenKeyset, adLockReadOnly
    txtbal.Text = portdb("cashbal")
    txtcurbal.Text = txtbal.Text
    portdb.Close
End If
End Sub
Private Sub cmdBsave_Click()
Dim stockdb As ADODB.Recordset
Dim compdb As ADODB.Recordset
Set stockdb = New ADODB.Recordset
stockdb.Open "select * from " & cmbBport.Text, scnn, adOpenKeyset, adLockOptimistic
    If stockdb.RecordCount > 0 Then stockdb.MoveFirst
    stockdb.Find "compcode = " & Trim(txtBcode.Text)
    If stockdb.EOF Then
        stockdb.AddNew
        stockdb("compcode") = Trim(txtBcode)
        stockdb("compname") = UCase(Trim(txtcoEname.Text))
        stockdb("stock") = CDbl(txtBqty.Text)
        stockdb("cost") = CDbl(txtBcost.Text)
        stockdb("avgcost") = stockdb("cost") / stockdb("stock")
        stockdb("mktprice") = stockdb("avgcost")
        stockdb("mktval") = stockdb("cost")
        stockdb("gain") = 0
        stockdb("gainper") = 0
        stockdb.Update
    Else
        stockdb("stock") = stockdb("stock") + CDbl(txtBqty.Text)
        stockdb("cost") = stockdb("cost") + CDbl(txtBcost.Text)
        stockdb("avgcost") = stockdb("cost") / stockdb("stock")
        stockdb("mktval") = stockdb("stock") * stockdb("mktprice")
        stockdb("gain") = stockdb("mktval") - stockdb("cost")
        stockdb("gainper") = (stockdb("gain") / stockdb("cost"))
        stockdb.Update
    End If
    Call updport(cmbBport.List(cmbBport.ListIndex, 1), (stockdb("mktprice") * CDbl(txtBqty.Text)), CDbl(txtBcost.Text), 1)
    Set compdb = New ADODB.Recordset
    compdb.Open "select * from companies where compcode = " & Val(txtBcode.Text), scnn, adOpenKeyset, adLockOptimistic
    If compdb.RecordCount <= 0 Then
        compdb.AddNew
        compdb("compcode") = Val(txtBcode.Text)
        compdb("compename") = UCase(Trim(txtcoEname.Text))
        compdb("companame") = Trim(txtcoAname.Text)
        compdb.Update
    End If
    compdb.Close
stockdb.Close
Unload frmtran
End Sub
Private Sub cmdcancel_Click(Index As Integer)
Unload frmtran
End Sub
Private Sub cmdSsave_Click()
Dim stockdb As ADODB.Recordset
Dim compdb As ADODB.Recordset
Set stockdb = New ADODB.Recordset
stockdb.Open "select * from " & cmbSport.Text, scnn, adOpenKeyset, adLockOptimistic
If stockdb.RecordCount > 0 Then
    stockdb.MoveFirst
    stockdb.Find "compcode = " & cmbcomp.List(cmbcomp.ListIndex, 1)
    If stockdb.EOF Then
        stockdb.AddNew
        stockdb("compcode") = cmbcomp.List(cmbcomp.ListIndex, 1)
        stockdb("compname") = UCase(Trim(cmbcomp.Text))
        stockdb("stock") = CDbl(txtsqty.Text) * -1
        stockdb("cost") = CDbl(txtsval.Text) * -1
        stockdb("avgcost") = stockdb("cost") / stockdb("stock")
        stockdb("mktprice") = stockdb("avgcost")
        stockdb("mktval") = CDbl(txtsval.Text) * -1
        stockdb("gain") = CDbl(txtsval.Text) * -1
        stockdb("gainper") = 0
        stockdb.Update
    Else
        stockdb("stock") = stockdb("stock") - CDbl(txtsqty.Text)
        stockdb("cost") = stockdb("cost") - (stockdb("avgcost") * CDbl(txtsqty.Text))
        stockdb("avgcost") = stockdb("cost") / stockdb("stock")
        stockdb("mktval") = stockdb("stock") * stockdb("mktprice")
        stockdb("gain") = stockdb("mktval") - stockdb("cost")
        stockdb("gainper") = (stockdb("gain") / stockdb("cost"))
        stockdb.Update
    End If
    Call updport(cmbSport.List(cmbSport.ListIndex, 1), (stockdb("mktprice") * CDbl(txtsqty.Text)), (stockdb("avgcost") * CDbl(txtsqty.Text)), 2)
    Set compdb = New ADODB.Recordset
    compdb.Open "select * from companies where compcode = " & cmbcomp.List(cmbcomp.ListIndex, 1), scnn, adOpenKeyset, adLockOptimistic
    If compdb.RecordCount <= 0 Then
        compdb.AddNew
        compdb("compcode") = cmbcomp.List(cmbcomp.ListIndex, 1)
        compdb("compename") = UCase(Trim(cmbcomp.Text))
        compdb("companame") = Trim(cmbcomp.Text)
        compdb.Update
    End If
    compdb.Close
End If
stockdb.Close
Unload frmtran
End Sub
Private Sub cmdWsave_Click()
Dim portdb As ADODB.Recordset
Set portdb = New ADODB.Recordset
portdb.Open "select * from portfolios where portfoliono = " & cmbWport.List(cmbWport.ListIndex, 1), scnn, adOpenKeyset, adLockOptimistic
If portdb.RecordCount > 0 Then
    If lsttype.ListIndex = 1 Then
        portdb("cashtran") = portdb("cashtran") - CDbl(txtWamt.Text)
        portdb("cashbal") = portdb("cashopen") + portdb("cashtran")
        portdb.Update
    ElseIf lsttype.ListIndex = 2 Then
        portdb("cashtran") = portdb("cashtran") + CDbl(txtWamt.Text)
        portdb("cashbal") = portdb("cashopen") + portdb("cashtran")
        portdb.Update
    End If
End If
portdb.Close
Unload frmtran
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Call fillcomb(Me.cmbBport, 0, 1)
Call fillcomb(Me.cmbSport, 0, 1)
Call fillcomb(Me.cmbWport, 0, 1)
Call fillcomb(Me.cmbcomp, 0, 2)
End Sub
Private Sub Form_Unload(Cancel As Integer)
Call frmcurrentport.updcomp
End Sub
Private Sub txtBcode_LostFocus()
DoEvents
If Trim(txtBcode.Text) <> "" Then
    Label17.Visible = True
    If checkcomp(Trim(txtBcode.Text)) = False Then
       txtBcode.Text = ""
       txtBcode.SetFocus
    End If
    Label17.Visible = False
End If
End Sub
Private Sub txtBcost_LostFocus()
If cmbBport.ListIndex > 0 Then
    If Trim(txtcoEname.Text) <> "" Then
        If Val(txtBqty.Text) <> 0 Then
            cmdBsave.Enabled = True
        End If
    End If
End If
End Sub
Sub updport(portno As Double, Amktval As Double, Acost As Double, updtype As Double)
Dim portmain As ADODB.Recordset
Set portmain = New ADODB.Recordset
portmain.Open "select * from portfolios where portfoliono = " & portno, scnn, adOpenKeyset, adLockOptimistic
If portmain.RecordCount > 0 Then
    With portmain
        If updtype = 1 Then ' buy tran
            .Fields("cashtran") = .Fields("cashtran") + Acost
            .Fields("cashbal") = .Fields("cashopen") - .Fields("cashtran")
            .Fields("cost") = .Fields("cost") + Acost
            .Fields("mktval") = .Fields("mktval") + Amktval
            .Fields("gain") = .Fields("mktval") - .Fields("cost")
            .Fields("gainper") = (.Fields("gain") / .Fields("cost"))
            .Update
        ElseIf updtype = 2 Then
            .Fields("cashtran") = .Fields("cashtran") - Acost
            .Fields("cashbal") = .Fields("cashopen") - .Fields("cashtran")
            .Fields("cost") = .Fields("cost") - Acost
            .Fields("mktval") = .Fields("mktval") - Amktval
            .Fields("gain") = .Fields("mktval") - .Fields("cost")
            .Fields("gainper") = (.Fields("gain") / .Fields("cost"))
            .Update
        ElseIf updtype = 3 Then 'withdraw/deposit cash
            If lsttype.ListIndex = 0 Then
                Acost = Acost * -1
            End If
            .Fields("cashtran") = .Fields("cashtran") - Acost
            .Fields("cashbal") = .Fields("cashopen") - .Fields("cashtran")
            .Update
        End If
    End With
End If
portmain.Close
End Sub
Private Sub txtsval_Change()
If Val(txtsqty.Text) > 0 Then
    cmdSsave.Enabled = True
Else
    cmdSsave.Enabled = False
End If
End Sub
Private Sub txtWamt_LostFocus()
If Trim(txtWamt.Text) <> "" Then
    If Val(txtWamt.Text) <> 0 Then
        If lsttype.ListIndex = 1 Then
            txtcurbal.Text = CDbl(txtbal.Text) - CDbl(txtWamt.Text)
        ElseIf lsttype.ListIndex = 2 Then
            txtcurbal.Text = CDbl(txtbal.Text) + CDbl(txtWamt.Text)
        End If
    End If
    cmdWsave.Enabled = True
Else
    cmdWsave.Enabled = False
End If
End Sub
