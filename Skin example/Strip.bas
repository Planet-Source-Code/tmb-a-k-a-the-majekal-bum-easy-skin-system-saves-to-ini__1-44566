Attribute VB_Name = "Strip"
Function StripFile(path) As String
On Error GoTo Bad
    Dim p%, cntr%
    StripPath$ = path
    p% = InStr(path, "\")


    Do While p%
        cntr% = p%
        p% = InStr(cntr% + 1, path, "\")
    Loop
    If cntr% > 0 Then StripPath$ = Mid$(path, cntr% + 1)
    path = Left(path, Len(path) - Len(StripPath$) - 1)
    
Bad:
End Function
