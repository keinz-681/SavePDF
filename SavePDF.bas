Attribute VB_Name = "SavePDF"
'made by keinz-681

Sub SavePDF()
Attribute SavePDF.VB_ProcData.VB_Invoke_Func = " \n14"
Dim name As String
Dim user As String
user = Environ("Username")
name = Format(Now, "yyyy-mm-dd-hh-nn-ss")
'MsgBox (name) 'Check file name
Dim fname As String
fname = "C:\Users\" + user + "\Desktop\" + name + ".pdf"
'MsgBox (fname) 'Check path name
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        fname, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        True 'Save  PDF and open soon
End Sub
