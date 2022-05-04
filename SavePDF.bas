Attribute VB_Name = "SavePDF"
'made by keinz-681

Sub SavePDF()
Attribute SavePDF.VB_ProcData.VB_Invoke_Func = " \n14"
Dim name As String
Dim user As String
user = Environ("Username")
name = Format(Now, "yyyy-mm-dd-hh-nn-ss")
'MsgBox (name) 'Check file name
Dim fname(3) As String
fname(0) = "C:\Users\"
fname(1) = "\Desktop\"
fname(2) = ".pdf"
fname(3) = fname(0) + user + fname(1) + name + fname(2)
'MsgBox (fname(3)) 'Check path name
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        fname(3), Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=True, OpenAfterPublish:= _
        True 'Save  PDF and open soon 
End Sub
