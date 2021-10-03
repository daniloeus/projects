Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

Sub email_range()

Dim OutApp As Object
Dim OutMail As Object
Dim count_rows, count_col As Integer
Dim pop As Range
Dim str1, str2 As String

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

count_rows = WorksheetFunction.CountA(Range("A1", Range("A1").End(xlDown)))
count_col = WorksheetFunction.CountA(Range("A1", Range("A1").End(xlToRight)))

Set pop = ActiveSheet.Range(Cells(1, 1), Cells(count_rows, count_col))

'str1 = "<BODY style = font-size: 11pt; font-family:Calibri>" & _
        "Boa tarde a todos,<br><br>Time,  temos as seguintes DMR's de Curitiba em aberto :<br>"
str1 = ActiveSheet.Range("L5").Value
'str2 = "<br>Por gentileza providenciar  a resolução e fechamento delas.<br>As DMR 's em vermelho estão abertas a mais de 60 dias =>  Por gentileza fechá-las o mais rápido possível.<br>Caso necessitem de algum suporte, nos avisem.<br><br>Atenciosamente"
str2 = ActiveSheet.Range("L14").Value

On Error Resume Next
    With OutMail
        .to = ActiveSheet.Range("L1").Value
        .CC = ActiveSheet.Range("L2").Value
        .BCC = ""
        .Subject = ActiveSheet.Range("L3").Value
        .Display
        .HTMLBody = str1 & RangetoHTML(pop) & str2 & .HTMLBody
    End With
    On Error GoTo 0
    
Set OutMail = Nothing
Set OutApp = Nothing

End Sub


Option Explicit
Dim lastrow As Integer
Sub convert2date()

Dim c As Range
Dim c2d As Range

Sheets("Sheet 1").Select

Application.ScreenUpdating = False
    For Each c In Range("B2", Range("B2").End(xlDown))
'        c.Value = DateValue(c.Value)
         c.Value = DateSerial(Right(c.Value, 4), Left(c.Value, 2), Mid(c.Value, 4, 2))

    Next
Application.ScreenUpdating = True

MsgBox ("Datas convertidas (atenção: executar a macro novamente causará erro nas datas)")

End Sub

Sub hidecols()
    
Columns("C:F").Hidden = True
Columns("H:I").Hidden = True
Columns("L:W").Hidden = True
Columns("Z").Hidden = True

End Sub
Sub unhidecols()

Columns("A:AA").Hidden = False

End Sub
Sub Filter_opens()
'filtrar  Issue Status em aberto
ActiveSheet.Range("A1:AA1" & lastrow).AutoFilter Field:=27, Criteria1:=Array("Issue Created", "Issue Updated", "Issue Reassigned") _
                                                                            , Operator:=xlFilterValues

End Sub

Sub Filter_CUR()
ActiveSheet.Range("A1:AA1" & lastrow).AutoFilter
ActiveSheet.Range("A1:AA1" & lastrow).AutoFilter Field:=10, Criteria1:=Array("Curitiba VW", "Curitiba") _
                                                                            , Operator:=xlFilterValues
Filter_opens
'Copy filtered table and paste it in Destination cell.
ActiveSheet.Range("A1:AA1" & lastrow).SpecialCells(xlCellTypeVisible).Copy
Sheets("CUR").Select
Range("A1").PasteSpecial xlPasteValues
Application.CutCopyMode = False
End Sub
Sub Filter_GVT()
ActiveSheet.Range("A1:AA1" & lastrow).AutoFilter
ActiveSheet.Range("A1:AA1" & lastrow).AutoFilter Field:=10, Criteria1:=Array("Gravatai", "Gravatai JIT", "Gravatai Foam") _
                                                                            , Operator:=xlFilterValues
Filter_opens
'Copy filtered table and paste it in Destination cell.
ActiveSheet.Range("A1:AA1" & lastrow).SpecialCells(xlCellTypeVisible).Copy
Sheets("GVT").Select
Range("A1").PasteSpecial xlPasteValues
Application.CutCopyMode = False

End Sub
Sub Filter_PAL()
ActiveSheet.Range("A1:AA1" & lastrow).AutoFilter
ActiveSheet.Range("A1:AA1" & lastrow).AutoFilter Field:=10, Criteria1:=Array("Pouso Alegre", "Pouso Alegre Foam", "Pouso Alegre Trim") _
                                                                            , Operator:=xlFilterValues
Filter_opens
'Copy filtered table and paste it in Destination cell.
ActiveSheet.Range("A1:AA1" & lastrow).SpecialCells(xlCellTypeVisible).Copy
Sheets("PAL").Select
Range("A1").PasteSpecial xlPasteValues
Application.CutCopyMode = False

End Sub
Sub Filter_ROS()
ActiveSheet.Range("A1:AA1" & lastrow).AutoFilter
ActiveSheet.Range("A1:AA1" & lastrow).AutoFilter Field:=10, Criteria1:=Array("Rosario", "Rosario Covers", "Rosario JIT", "Rosario Trim" _
                                                                            , "Argentina Covers", "Argentina JIT"), Operator:=xlFilterValues
Filter_opens
'Copy filtered table and paste it in Destination cell.
ActiveSheet.Range("A1:AA1" & lastrow).SpecialCells(xlCellTypeVisible).Copy
Sheets("ROS").Select
Range("A1").PasteSpecial xlPasteValues
Application.CutCopyMode = False

End Sub
Sub Filter_SBC()
ActiveSheet.Range("A1:AA1" & lastrow).AutoFilter
ActiveSheet.Range("A1:AA1" & lastrow).AutoFilter Field:=10, Criteria1:=Array("Sao Bernardo JIT", "Sao Bernardo do Campos JIT", "Sao Bernardo do Campo JIT", "Sao Bernardo Ford" _
                                                                            , "Sao Bernardo Interiors", "Sao Bernardo do Campo Interiors" _
                                                                            , "Sao Bernardo Foam", "Sao Bernardo do Campo Foam") _
                                                                            , Operator:=xlFilterValues
Filter_opens
'Copy filtered table and paste it in Destination cell.
ActiveSheet.Range("A1:AA1" & lastrow).SpecialCells(xlCellTypeVisible).Copy
Sheets("SBC").Select
Range("A1").PasteSpecial xlPasteValues
Application.CutCopyMode = False

End Sub

Sub goBack_Clearfilter()
'Back to Sheet 1
Sheets("Sheet 1").Select
'Remove filter that was applied.
ActiveSheet.AutoFilterMode = False
End Sub

Sub Generate_reports()

'define last row with data to filter
lastrow = Sheet1.Cells(Rows.Count, 1).End(xlUp).Row

'ensure to select "Sheet 1"
Sheets("Sheet 1").Select
hidecols 'only cols need

'Update CUR
Filter_CUR
goBack_Clearfilter

'Update GVT
Filter_GVT
goBack_Clearfilter

'Update PAL
Filter_PAL
goBack_Clearfilter

'Update ROS
Filter_ROS
goBack_Clearfilter

'Update SBC
Filter_SBC
goBack_Clearfilter

'Unhide cols (get DF back to original view)
unhidecols

MsgBox ("Relatórios gerados para cada planta com Sucesso!!")


End Sub
