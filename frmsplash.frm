VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmsplash 
   BackColor       =   &H00000000&
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6390
   ControlBox      =   0   'False
   Icon            =   "frmsplash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   2280
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   120
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   960
      Picture         =   "frmsplash.frx":08CA
      Top             =   1200
      Width           =   4500
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub checkmkt()
Dim thepage As String, returnstr As String, exactstr As String
Dim beg As Integer, endline As Integer
On Error Resume Next
thepage = "http://www.tadawul.com.sa/changelang.asp?language=en"
thepage = Inet1.OpenURL(thepage, icString)
returnstr = Inet1.GetChunk(2048, icString)
Do While Len(returnstr) <> 0
    DoEvents
    returnstr = Inet1.GetChunk(2048, icString)
Loop


DoEvents
Label1.Caption = "Checking market status...."
thepage = "http://www.tadawul.com.sa/user/myhome.asp"
returnstr = thepage
thepage = Inet1.OpenURL(thepage, icString)
returnstr = Inet1.GetChunk(2048, icString)
Do While Len(returnstr) <> 0
    DoEvents
    returnstr = Inet1.GetChunk(2048, icString)
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
Label1.Caption = "connecting to database ..."
Set scnn = New ADODB.Connection
scnn.Provider = "Microsoft.Jet.OLEDB.3.51"
scnn.Open App.Path & "\tadawul.mdb", "admin", ""
Unload frmsplash
MDIForm1.Show
End Sub
Private Sub Timer1_Timer()
Timer1.Interval = 0
Call checkmkt
End Sub
