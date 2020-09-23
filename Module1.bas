Attribute VB_Name = "Module1"
Public griddb As ADOR.Recordset
Public scnn As ADODB.Connection
Public mktstatus As String, mktopen As Boolean
Sub fillcomb(varobject As Object, fillzero As Double, filltype As Double)
' fillzero  : to fill the zero index with "all" , 1 = "All" ; 2 = " "
' filltype : 1 for portfolio, 2 for company

Dim comblist As ADODB.Recordset, sql As String, t
Set comblist = New ADODB.Recordset
If filltype = 1 Then
    sql = "select portfoliono as Vnumb,portfolioname as Vname from portfolios order by portfoliono"
ElseIf filltype = 2 Then
    sql = "select compcode as Vnumb, compename as Vname from companies order by compcode"
End If
comblist.Open sql, scnn, adOpenKeyset, adLockReadOnly
t = 0
If comblist.RecordCount > 0 Then
    comblist.MoveFirst
    varobject.Clear
    If fillzero = 1 Then
        varobject.AddItem "All"
        varobject.List(t, 1) = t
    Else
        varobject.AddItem " "
        varobject.List(t, 1) = ""
    End If
    Do While Not comblist.EOF
        t = t + 1
        With varobject
            .AddItem IIf(IsNull(comblist("Vname")), " ", Trim(comblist("Vname")))
            .List(t, 1) = comblist("vnumb")
        End With
        comblist.MoveNext
    Loop
    varobject.ListIndex = 1
End If
comblist.Close
End Sub

