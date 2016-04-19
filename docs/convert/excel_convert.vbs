Set app = CreateObject("Excel.Application")
Set fso = CreateObject("Scripting.FileSystemObject")
strPath = "C:\work\repositorios\fist\docs\convert\"
Set fld = fso.GetFolder(strPath)
For Each fil In fld.Files
    If Right(fil.Name, 3) = "xls" Then
        Set wbk = app.Workbooks.Open(fil)
        If wbk.HasVBProject Then
            wbk.SaveAs fil & "m", 52
        Else
            wbk.SaveAs fil & "x", 51
        End If
        wbk.Close False
    End If
Next
app.Quit

