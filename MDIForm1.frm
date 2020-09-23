VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H0080C0FF&
   Caption         =   "Tadawul Stocks"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7815
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "MDIForm1.frx":08CA
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin InetCtlsObjects.Inet inet 
      Left            =   120
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   65535
      Left            =   480
      Top             =   7080
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "20/01/2004"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "04:30 PM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuport 
      Caption         =   "Portfolio"
      Begin VB.Menu mnuaddport 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuview 
         Caption         =   "View"
      End
      Begin VB.Menu mnudelport 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tempdb As ADOR.Recordset, timercnt As Integer
Private Sub MDIForm_Load()
timercnt = 0
MDIForm1.Caption = "Saudi Stocks ...Connected...." & mktstatus
End Sub
Sub updprices()
Dim i As Integer, thepage As String, returnstr As String, correctstr As String
Dim beg As Double, endline As Double
Dim compname As String, price As String
Dim cloprice As String, highprice As String, lowprice As String
Dim compdb As ADODB.Recordset

On Error GoTo errh:
StatusBar1.Panels(6).Text = "checking for english language.."
thepage = "http://www.tadawul.com.sa/changelang.asp?language=en"
thepage = inet.OpenURL(thepage, icString)
returnstr = inet.GetChunk(2048, icString)
Do While Len(returnstr) <> 0
    DoEvents
    returnstr = inet.GetChunk(2048, icString)
Loop

Set compdb = New ADODB.Recordset
compdb.Open "select * from companies order by compcode", scnn, adOpenKeyset, adLockReadOnly
If compdb.RecordCount > 0 Then
    compdb.MoveFirst
    Set tempdb = New ADOR.Recordset
    tempdb.Fields.Append "compcode", adDouble
    tempdb.Fields.Append "price", adDouble
    tempdb.Open
    Do While Not compdb.EOF
        tempdb.AddNew
        tempdb("compcode") = compdb("compcode")
        tempdb("price") = 0
        tempdb.Update
        DoEvents
        StatusBar1.Panels(6).Text = "checking price for.." & compdb("compcode")
        thepage = "http://www.tadawul.com.sa/quotes/quote.asp?QuoteCode=" & compdb("compcode")
        returnstr = thepage
        thepage = inet.OpenURL(thepage, icString)
        returnstr = inet.GetChunk(2048, icString)
        Do While Len(returnstr) <> 0
            DoEvents
            returnstr = inet.GetChunk(2048, icString)
        Loop
        beg = 0
        beg = InStr(1, thepage, "Volume")
        If beg <> 0 Then
            beg = InStr(beg, thepage, "Number")
            If beg <> 0 Then
                beg = InStr(beg, thepage, ">")
                If beg <> 0 Then
                    beg = beg + 1
                    endline = InStr(beg, thepage, "<")
                    If endline <> 0 Then
                        correctstr = Mid(thepage, beg, (endline - beg))
                        price = correctstr
                    End If
                    beg = beg + Len(correctstr)
                End If
            End If
        End If
        If IsNumeric(price) Then
            tempdb("price") = CDbl(price)
            tempdb.Update
        End If
    compdb.MoveNext
    Loop
Call updport
End If
Exit Sub

errh:
On Error Resume Next
End Sub
Sub updport()
Dim portmain As ADODB.Recordset, Tsql As String
Set portmain = New ADODB.Recordset
portmain.Open "select * from portfolios", scnn, adOpenKeyset, adLockReadOnly
StatusBar1.Panels(6).Text = "Updating portfolio.."
If portmain.RecordCount > 0 Then
    If tempdb.RecordCount > 0 Then
        tempdb.MoveFirst
        Do While Not tempdb.EOF
            If tempdb("price") <> 0 Then
                portmain.MoveFirst
                Do While Not portmain.EOF
                    Tsql = ""
                    Tsql = "update " & portmain("portfolioname") & _
                    " set mktprice = " & tempdb("price") & _
                    " where compcode = " & tempdb("compcode")
                    scnn.Execute Tsql
                    
                    Tsql = "update " & portmain("portfolioname") & _
                    " set mktval = mktprice * stock " & _
                    " where compcode = " & tempdb("compcode")
                    scnn.Execute Tsql

                    Tsql = "update " & portmain("portfolioname") & _
                    " set gain = mktval - cost " & _
                    " where compcode = " & tempdb("compcode")
                    scnn.Execute Tsql

                    Tsql = "update " & portmain("portfolioname") & _
                    " set gainper = gain/cost " & _
                    " where compcode = " & tempdb("compcode")
                    scnn.Execute Tsql

                    portmain.MoveNext
                Loop
            End If
            tempdb.MoveNext
        Loop
    End If
Call updportdtls
End If
portmain.Close
StatusBar1.Panels(6).Text = "Updated.."
End Sub
Private Sub mnuaddport_Click()
frmcreate.Show
End Sub

Private Sub mnudelport_Click()
frmdel.Show
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuview_Click()
frmcurrentport.Show
End Sub
Private Sub Timer1_Timer()
timercnt = timercnt + 1
StatusBar1.Panels(6).Text = "Next update after " & (2 - timercnt) & " mnt(s)..."
If timercnt = 2 Then
    Timer1.Interval = 0
    Call checkmkt
    If mktopen Then
        Call updprices
    End If
    Timer1.Interval = 65535
    timercnt = 0
End If
End Sub
Sub updportdtls()
Dim portdb As ADODB.Recordset, Psql As String
Dim stockdb As ADODB.Recordset
Set portdb = New ADODB.Recordset
portdb.Open "select * from portfolios order by portfoliono", scnn, adOpenKeyset, adLockOptimistic
If portdb.RecordCount > 0 Then
    portdb.MoveFirst
    Do While Not portdb.EOF
        Psql = "select sum(cost) as Tcost,sum(mktval) as Tmkt " & _
            " from " & portdb("portfolioname")
        Set stockdb = New ADODB.Recordset
        stockdb.Open Psql, scnn, adOpenKeyset, adLockReadOnly
        If stockdb.RecordCount > 0 Then
            portdb("cost") = stockdb("tcost")
            portdb("mktval") = stockdb("tmkt")
            portdb("gain") = stockdb("tmkt") - stockdb("tcost")
            portdb("gainper") = Round((portdb("gain") / portdb("cost")), 2)
            portdb.Update
        End If
        portdb.MoveNext
    Loop
End If
portdb.Close
End Sub
Sub checkmkt()
Dim thepage As String, returnstr As String, exactstr As String
Dim beg As Integer, endline As Integer

On Error GoTo errh:
thepage = "http://www.tadawul.com.sa/changelang.asp?language=en"
thepage = inet.OpenURL(thepage, icString)
returnstr = inet.GetChunk(2048, icString)
Do While Len(returnstr) <> 0
    DoEvents
    returnstr = inet.GetChunk(2048, icString)
Loop


DoEvents
StatusBar1.Panels(6).Text = "Checking market status...."
thepage = "http://www.tadawul.com.sa/user/myhome.asp"
returnstr = thepage
thepage = inet.OpenURL(thepage, icString)
returnstr = inet.GetChunk(2048, icString)
Do While Len(returnstr) <> 0
    DoEvents
    returnstr = inet.GetChunk(2048, icString)
Loop
beg = 0
beg = InStr(1, thepage, "changeLang")
mktstatus = ""
If beg > 0 Then
    beg = InStr(beg, thepage, "COLOR")
    If beg > 0 Then
        beg = InStr(beg, thepage, ">")
        beg = beg + 1
        endline = InStr(beg, thepage, "<")
        If endline > 0 Then
            exactstr = Mid(thepage, beg, (endline - beg))
            mktstatus = Trim(exactstr)
            MDIForm1.Caption = "Saudi Stocks ...Connected...." & mktstatus
            StatusBar1.Panels(7).Text = mktstatus
            If InStr(1, UCase(mktstatus), UCase("pre open")) > 0 Then
                mktopen = False
            ElseIf InStr(1, UCase(mktstatus), UCase("open")) > 0 Then
                mktopen = True
            Else
                mktopen = False
            End If
        Else
            mktstatus = "error....."
        End If
    Else
        mktstatus = "error....."
    End If
Else
    mktstatus = "error..."
End If
StatusBar1.Panels(6).Text = ""
Exit Sub

errh:
mktstatus = "error..."
End Sub
