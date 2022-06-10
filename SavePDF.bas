Attribute VB_Name = "SavePDF"
'made by keinz-681

Sub SavePDF()
Dim fname As String
'MsgBox (fname) 'Check path name
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:="C:\Users\" + Environ("Username") + "\Desktop\" + Format(Now, "yyyy-mm-dd-hh-nn-ss") + ".pdf" _
        , Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        True 'Save  PDF and open soon
End Sub
