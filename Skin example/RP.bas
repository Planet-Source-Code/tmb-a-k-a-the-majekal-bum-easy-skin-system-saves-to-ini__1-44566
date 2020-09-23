Attribute VB_Name = "RP"
Public Function RemovePath(ByVal path As String)
'When called upon this removes everything before the file name in the path.
    Do While InStr(path, "\") <> 0
        path = Right(path, Len(path) - InStr(path, "\"))
    Loop
    RemovePath = path
End Function
