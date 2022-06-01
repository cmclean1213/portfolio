Option Explicit
Public rg As Range
Public LC As String
Public WatchEmails As New Collection
Public OutApp As Outlook.Application
'Version 3.0
'Cody McLean
'12.10.2021
Public Sub VendReports3()
'1. Check if proper file, set report range
    speedToggle
    Set rg = Range("A1").CurrentRegion

'2. Location setup -> assign location #s to cribs & return main machine if machines > 1, exit if nothing
    If Cells(2, 3).Value <> "" Then LC = getLocation(Cells(2, 3).Value) Else Exit Sub
    
'3. Restructure report
    Restructure rg

'4. Main report styles
    ReportStyles LC, rg
    
'5. Assign data for emailing
    MailOut LC
    
'6. Exit out
    speedToggle
End Sub

Function speedToggle()
    If Application.ScreenUpdating = True Then
        If Cells(1, 1).Value <> "Bin Qty" And Cells(1, 8).Value <> "Serial ID" Then
            MsgBox "Report not supported, please use it on raw data of vend reports only!"
            Exit Function
        ElseIf Cells(2, 1).Value = "" Then
            Exit Function
        End If
    
        With Application
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
            .DisplayAlerts = False
            .DisplayStatusBar = False
            .EnableEvents = False
        End With
        With ActiveWindow
            .WindowState = xlNormal
            .Left = 1200
            .WindowState = xlMaximized
        End With
    
    Else
        With Application
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .DisplayAlerts = True
            .DisplayStatusBar = True
            .EnableEvents = True
        End With
    End If
End Function
Function getLocation(ByVal LC As String) As String
    getLocation = Left(LC, 2)
    Select Case getLocation
        Case 11 'Marwood 5
            getLocation = 6
        Case 21 'Schukra
            getLocation = 20
        Case 67, 68 'Rollstar
            getLocation = 66
        Case 41 'FECT
            getLocation = 40
        Case 61, 62, 63 'Alfield
            getLocation = 60
        Case 75 'FNG Peterborough
            getLocation = 74
        Case 81, 82 'MRE Tillsonburg
            getLocation = 80
    End Select
End Function
Function Restructure(rg As Range)
Dim i, j        As Integer

    Range("C:C,D:D,H:H,J:L,O:P").Delete
    Columns("G:H").Cut
    Columns("C:D").Insert shift:=xlToRight
    [A1:H1] = [{"Date", "Time","ID","Name","Item","Description","Qty","Price"}]
    
    For i = rg.Rows.Count To 1 Step -1
        If Range("A" & i).Value = "Grand Total" Or Range("A" & i).Value = "Notes:" Then Rows(i).Delete
        If LCase$(Range("D" & i).Value) Like "*tytan*" Or LCase$(Range("D" & i).Value) Like "*admin*" Or LCase$(Range("D" & i).Value) Like "*gabe*" _
        Or LCase$(Range("D" & i).Value) Like "*metro*" Then Rows(i).Delete
    Next i
       
    With Range("A1:H1")
        .Font.Bold = True
        .Interior.Color = 65535
    End With
    
    Range("A:A").NumberFormat = "m/d/yy"
    Range("B:B").NumberFormat = "[$-en-US]h:mm AM/PM;@"
    
    For i = 2 To rg.Rows.Count
        Cells(i, 2) = Left(Cells(i, 2), 5) & " " & Right(Cells(i, 2), Len(Cells(i, 2)) - 5)
    Next i
End Function
Function ReportStyles(ByVal LC As Integer, ByRef rg As Range)
    Dim refFile
    Dim r As Integer
    Dim i As Variant
    Dim glasses As Boolean
    
    Set rg = Range("A1").CurrentRegion
    
    With ActiveSheet.Sort
        .SortFields.Add Key:=Range("A2"), Order:=xlAscending
        .SortFields.Add Key:=Range("B2"), Order:=xlAscending
        .SetRange rg
        .Header = xlYes
        .Apply
    End With
            
    'PivotTable Summary Reports
    Select Case LC
        Case 2, 3, 4, 5, 6, 11, 15, 54, 77, 78, 79, 84
            ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rg).CreatePivotTable TableDestination:=Range("J3"), TableName:="SummaryPivot"
            With ActiveSheet.PivotTables("SummaryPivot")
                .ColumnGrand = True
                .SaveData = True
                .ShowValuesRow = False
                .RowAxisLayout xlTabularRow
                .ShowDrillIndicators = False
                With .PivotFields("Item")
                    .Orientation = xlRowField
                    .Subtotals(1) = False
                End With
                .PivotFields("Description").Orientation = xlRowField
                .ShowTableStyleRowHeaders = False
                .AddDataField ActiveSheet.PivotTables("SummaryPivot").PivotFields("Qty"), "Quantity", xlSum
                .AddDataField ActiveSheet.PivotTables("SummaryPivot").PivotFields("Price"), "Spend", xlSum
            End With
            Range("M:M").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            For Each i In ActiveSheet.PivotTables("SummaryPivot").PivotFields("Item").PivotItems
                If i.name = "TR-CRBK110" Then glasses = True
            Next i
            'TRW Glasses Count
            If LC = 54 And glasses = True Then
                ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rg).CreatePivotTable TableDestination:=Range("J" & Range("J" & Rows.Count).End(xlUp).Row + 2), TableName:="GlassesPivot"
                With ActiveSheet.PivotTables("GlassesPivot")
                    .ColumnGrand = True
                    .SaveData = True
                    .ShowValuesRow = False
                    .RowAxisLayout xlTabularRow
                    .ShowDrillIndicators = False
                    With .PivotFields("Item")
                        .Orientation = xlRowField
                        .Subtotals(1) = False
                        .ClearAllFilters
                        For Each i In .PivotItems
                            If i.name <> "TR-CRBK110" Then i.Visible = False
                        Next i
                    End With
                    With .PivotFields("Name")
                        .Orientation = xlRowField
                        .Subtotals(1) = False
                    End With
                    .AddDataField ActiveSheet.PivotTables("GlassesPivot").PivotFields("Qty"), "Quantity", xlSum
                    .PivotFields("Date").Orientation = xlRowField
                End With
            End If
        
        'Additional Columns for Ref Files
        Case 8, 60, 74, 80, 83
            'refFiles
            If LC = 8 Or LC = 74 Or LC = 80 Then
                Set refFile = GetObject("T:\CUSTOMER CENTER\CUSTOMERS\VENDING & VMI\Current Vending Accounts Data\CM to AX Items.xlsx").Sheets("Sheet1").Range("A1").CurrentRegion
            ElseIf LC = 60 Then
                Set refFile = GetObject("T:\CUSTOMER CENTER\CUSTOMERS\VENDING & VMI\Current Vending Accounts Data\Martinrea\Alfield\Employees\Alfield Departmental Employee List.xlsx").Worksheets("Master List").Range("A1").CurrentRegion
            ElseIf LC = 83 Then
                Set refFile = GetObject("T:\CUSTOMER CENTER\CUSTOMERS\VENDING & VMI\Account Limits\Moore Packaging\Moore Packaging Updated User List for Monday reports.xlsx").Worksheets("Master List").Range("A1").CurrentRegion
            End If
            With refFile
                'FNG/Ventra
                If LC = 8 Or LC = 74 Then
                    Range("F:F").Insert
                    Range("F1").Value = "Ventra Item"
                    For r = 2 To Range("A" & Rows.Count).End(xlUp).Row
                        Range("F" & r).Value = Application.VLookup(Range("E" & r), refFile, 5, 0)
                    Next r
                    If LC = 8 Then
                        With ActiveSheet.Sort
                            .SortFields.Clear
                            .SetRange rg
                            .SortFields.Add2 Key:=Range("D2"), SortOn:=xlSortOnValues, Order:=xlAscending
                            .Header = xlYes
                            .Orientation = xlTopToBottom
                            .Apply
                        End With
                    End If
                'Departmental Accounts (Alfield & Moore)
                ElseIf LC = 60 Or LC = 83 Then
                    Range("E:E").Insert
                    Range("E1").Value = "Department"
                    For r = 2 To Range("A" & Rows.Count).End(xlUp).Row
                        Range("E" & r).Value = Application.VLookup(Range("C" & r), refFile, 6, 0)
                        If IsError(Range("E" & r).Value) Then Range("E" & r).Value = ""
                    Next r
                'Martinrea Tillsonburg Custom
                ElseIf LC = 80 Then
                    Range("F:G").Insert
                    [F1:G1] = [{"Martinrea Vending Item", "Martinrea Item"}]
                    For r = 2 To Range("A" & Rows.Count).End(xlUp).Row
                        Range("F" & r).Value = Application.VLookup(Range("E" & r), refFile, 5, 0)
                        Range("G" & r).Value = Application.VLookup(Range("E" & r), refFile, 6, 0)
                    Next r
                End If
            End With
        End Select
End Function
Function MailOut(ByVal LC As Integer)
    Dim strSigFilePath, enviro, strBuffer, sender, msg, peterSendList, wbName, LM, LLM, TS, YD, LSat, LSun, lastMonth, rpath, rformat As String
    Dim olMail, objSignatureFile, objFSO As Object
    Dim ExApp As Excel.Application
    Dim thisMail As New EmailWatcher
    
    If OutApp Is Nothing Then Set OutApp = CreateObject("Outlook.Application")
    If ExApp Is Nothing Then Set ExApp = CreateObject("Excel.Application")
    
    enviro = CStr(Environ("appdata"))
    strSigFilePath = enviro & "\Microsoft\Signatures\"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objSignatureFile = objFSO.OpenTextFile(strSigFilePath & "Vending.htm")
    strBuffer = objSignatureFile.ReadAll
    objSignatureFile.Close
    
    LM = Format(DateAdd("ww", -1, Date - (Weekday(Date, vbMonday) - 1)), "MMM d")
    LLM = Format(DateAdd("ww", -2, Date - (Weekday(Date, vbMonday) - 1)), "MMM d")
    TS = Format(Date - Weekday(Date, vbSunday) + 1, "MMM d")
    YD = Format(Range("A2").Value, "MMM d")
    LSat = Format(DateAdd("ww", 0, Date - (Weekday(Date, vbSaturday) - 1)), "MMM d")
    LSun = Format(DateAdd("ww", -1, Date - (Weekday(Date, vbSunday) - 1)), "MMM d")
    With Worksheets("Sheet1")
        .name = LM & " - " & TS
        If LC = 40 Then .name = YD
        If LC = 60 Then .name = LSun & " - " & LSat
    End With
    
    If LC = 74 Then peterSendList = "akernachan@flexngate.com; dbell@flexngate.com; lhaylor@flexngate.com; jmoore2@flexngate.com; rwalmsley@flexngate.com; lbrown@flexngate.com; dgodfrey@flexngate.com; gholloway@flexngate.com; mschafer@flexngate.com; cwootton@flexngate.com;" & _
        "ppoole@flexngate.com; rweiss@flexngate.com; dhughes@flexngate.com; srocha@flexngate.com; hseaboyer@flexngate.com; ldaniels@flexngate.com; amartin@flexngate.com; kcrebar@flexngate.com; ryaworski@flexngate.com; achilcott@flexngate.com; ko'neill@flexngate.com;" & _
        "rhawryszko@flexngate.com; sthompson@flexngate.com; swhite@flexngate.com; jwelsh@flexngate.com; jmackay@flexngate.com; jwalsh@flexngate.com; vgonsalves@flexngate.com; sjackson@flexngate.com; jnodell@flexngate.com; bbergeron@flexngate.com; gmorrison@flexngate.com;" & _
        "rbakelaar@flexngate.com; mbateman@flexngate.com; tcurtin@flexngate.com; mdacruz@flexngate.com; jkellogg@flexngate.com; awiles@flexngate.com; dmccabe@flexngate.com; clindsay@flexngate.com; tzdunic@flexngate.com; tmorgan@flexngate.com; ljamieson@flexngate.com; kcourneyea@flexngate.com"
    
    msg = "<p>Good morning,</p> <p></p> <p>Please find attached the PPE vending report for the week of " & LM & "." & "<p></p> <p>Thank you and have a great week!</p> <p></p>" & strBuffer
    
    Cells.EntireColumn.AutoFit
    rpath = "\\TYTAN-DC\Company\CUSTOMER CENTER\CUSTOMERS\VENDING & VMI\Current Vending Accounts Data\"
    rformat = LM & " - " & TS & ".xlsx"
    
    Select Case LC
        Case 2 'Marwood 1
            ActiveWorkbook.SaveAs Filename:=rpath & "Marwood\Usage Reports\Plant 1\2020\Marwood Plant 1 Weekly Vend Report " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "bob.halasz@marwoodinternational.com;April.Parslow@marwoodinternational.com; Kelsie.Bouck@marwoodinternational.com"
        Case 3 'Marwood 2
            ActiveWorkbook.SaveAs Filename:=rpath & "Marwood\Usage Reports\Plant 2\2020\Marwood Plant 2 Weekly Vend Report " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "leo.chabot@marwoodinternational.com;bob.halasz@marwoodinternational.com;April.Parslow@marwoodinternational.com; Michael.Oakes@marwoodinternational.com "
        Case 4 'Marwood 3
            ActiveWorkbook.SaveAs Filename:=rpath & "Marwood\Usage Reports\Plant 3\2020\Marwood Plant 3 Weekly Vend Report " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "steve.saris@marwoodinternational.com;bob.halasz@marwoodinternational.com;April.Parslow@marwoodinternational.com; Michael.Oakes@marwoodinternational.com"
        Case 5 ' Marwood 4
            ActiveWorkbook.SaveAs Filename:=rpath & "Marwood\Usage Reports\Plant 4\2020\Marwood Plant 4 Weekly Vend Report " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "chris.long@marwoodinternational.com;bob.halasz@marwoodinternational.com;April.Parslow@marwoodinternational.com"
        Case 6 'Marwood 5
            ActiveWorkbook.SaveAs Filename:=rpath & "Marwood\Usage Reports\Plant 5\2020\Marwood Plant 5 Weekly Vend Report " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "steve.Demeyere@marwoodinternational.com; bob.halasz@marwoodinternational.com;April.Parslow@marwoodinternational.com"
        Case 7 'Aluma Bolton
            ActiveWorkbook.SaveAs Filename:=rpath & "Aluma Bolton\2020 Usage\Aluma Bolton Vending Usage " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "rhamilton@aluma.com"
        Case 8 'Ventra Plastics Windsor
            ActiveWorkbook.SaveAs Filename:=rpath & "Flex-N-Gate\Windsor\Vending Reports\Ventra Plastics Windsor Vend Usage Report " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "TNassar@flexngate.com; JDiab@flexngate.com; jmthomas@flexngate.com"
        Case 9 ' HydraDyne
            ActiveWorkbook.SaveAs Filename:=rpath & "HydraDyne\Usage Reports " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "Phil Gerantonis <philip.gerantonis@hydradynetech.com>"
        Case 10 'MIC
            ActiveWorkbook.SaveAs Filename:=rpath & "Martinrea\MIC\Vend Reports\MIC Vend Usage Report " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "krishan.kumar@martinrea.com"
        Case 12 'Agway
            ActiveWorkbook.SaveAs Filename:=rpath & "Agway Metals\Vend Reports\Agway Metals Vend Usage Report " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "dvitello@agwaymetals.com;mshouldice@agwaymetals.com;cvanmierlo@agwaymetals.com"
        Case 15 'Nova Steel
            ActiveWorkbook.SaveAs Filename:=rpath & "Nova Steel Stoney Creek\Vending Reports\2020\Nova Steel SC Weekly PPE Vend Usage " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "Bryson.Sherriff@novasteel.ca; jeff.robinson@novasteel.ca"
        Case 16 'Howard
            ActiveWorkbook.SaveAs Filename:=rpath & "Flex-N-Gate\Howard\Vend Reports\FNG Howard Vend Usage Report " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "khart@flexngate.com"
        Case 20 'Schukra
            ActiveWorkbook.SaveAs Filename:=rpath & "Schukra Vending\Usage Reports\2020\Schukra Usage Report " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "lonnie.brown@leggett.com;mingwei.sun@leggett.com"
        Case 22 'JMP
            ActiveWorkbook.SaveAs Filename:=rpath & "JMP\Vending Reports\JMP Weekly PPE Vend Usage " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "vpupatello@vintechstampings.com; ap@jmpmetals.com; production@vintechstampings.com"
        Case 23 'Mahle
            ActiveWorkbook.SaveAs Filename:=rpath & "Mahle\Vending Reports\Mahle Weekly PPE Vend Usage " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "brandon.lindquist@mahle.com"
        Case 24 ' VAW
            ActiveWorkbook.SaveAs Filename:=rpath & "VAW\Vending Reports\VAW Weekly PPE Vend Usage " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "SRuggaber@flexngate.com;mfurtado@flexngate.com"
        Case 40 'FECT
            'Daily Report
            If ActiveWorkbook.name Like "FECT Daily Usage*" Then
                ActiveWorkbook.SaveAs Filename:=rpath & "Faurecia Brampton Vending\Reports\2020\Daily Reports\FECT Daily Usage Report - " & YD & ".xlsx", FileFormat:=xlOpenXMLWorkbook
                sender = "alex.duica@faurecia.com;business.admin@forvia.com;michael.corpuz@forvia.com"
                msg = "<p>Good morning everyone,</p> <p></p> <p>Please see attached for yesterday's daily usage for " & YD & "<p></p> <p>Thank you and have a great day!</p> <p></p>" & strBuffer
            'Weekly Report
            Else
                ActiveWorkbook.SaveAs Filename:=rpath & "Faurecia Brampton Vending\Reports\2020\Vending Reports\FECT Vending Report " & rformat, FileFormat:=xlOpenXMLWorkbook
                sender = "business.admin@forvia.com;alex.duica@faurecia.com;michael.corpuz@forvia.com"
            End If
        Case 52 'Martinrea Ingersoll
            ActiveWorkbook.SaveAs Filename:=rpath & "Martinrea\Ingersoll\2020 Usage\Ingersoll Vending Usage " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "lisa.janssen@martinrea.com; Bob.Ashby@martinrea.com"
        Case 54 'TRW
            ActiveWorkbook.SaveAs Filename:=rpath & "TRW Windsor\Cribmaster Reports\TRW Weekly Vend Usage Report " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "amit.bumma@zf.com"
        Case 60 'Alfield Departmental Daily Usage
            ActiveWorkbook.SaveAs Filename:=rpath & "Martinrea\Alfield\2020\Alfield Weekly Departmental Usage - " & LSun & " - " & LSat & ".xlsx", FileFormat:=xlOpenXMLWorkbook
            sender = "peter.gismondi@martinrea.com;skanda.pathmanathan@martinrea.com;Debbie.guo@martinrea.com;nasreen.zuberi@martinrea.com"
            msg = "<p>Good morning everyone,</p> <p></p> <p>Please find attached the PPE departmental usage vending report for the week of " & LSun & " - " & LSat & "<p></p> <p>Thank you and have a great day!</p> <p></p>" & strBuffer
        Case 64 'Martinrea Dresden
            ActiveWorkbook.SaveAs Filename:=rpath & "Martinrea\Dresden\Usage Reports\2020\Martinrea Dresden Detailed Vend Usage Report from Tytan " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = ""
        Case 66 'Rollstar
            ActiveWorkbook.SaveAs Filename:=rpath & "Martinrea\Rollstar\2020\Rollstar Detailed Usage Report from Tytan " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "pjohnson@martinrea.com"
        Case 69 'FNG Tottenham
            ActiveWorkbook.SaveAs Filename:=rpath & "Flex-N-Gate\Tottenham\Vend Usage Report\FNG Tottenham Vend Usage Report " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "NDearsley@flexngate.com"
        Case 74 'Ventra Plastics Peterborough
            ActiveWorkbook.SaveAs Filename:=rpath & "Flex-N-Gate\Peterborough\Vending Reports\Flex-N-Gate Peterborough Vend Usage Report " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = peterSendList
        Case 76 'Flex Canada
            ActiveWorkbook.SaveAs Filename:=rpath & "Flex-N-Gate\Canada\Vending Reports\Flex Canada Vend Usage Report " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "MBaioff@flexngate.com;DPatry@flexngate.com;PRusso@flexngate.com"
        Case 78 'Team Sarnia
            ActiveWorkbook.SaveAs Filename:=rpath & "Team Industries (Sarnia)\Vend Usage Report\Team Sarnia Vend Usage Report " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "steve.fisher@TeamInc.com;rob.moulton@teaminc.com;conrad.vandermeulen@teaminc.com"
            msg = "<p>Good morning,</p> <p></p> <p>Please find attached the PPE vending report for the period of " & LM & " - " & TS & "<p></p> <p>Thank you and have a great week!</p> <p></p>" & strBuffer
        Case 79 'Team Sarnia Plank Rd
            ActiveWorkbook.SaveAs Filename:=rpath & "Team Sarnia (Plank)\Vend Usage Report\Team Sarnia Plank Rd Vend Usage Report " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "stu.poderys@teaminc.com"
        Case 80 'Martinrea Tillsonburg
            ActiveWorkbook.SaveAs Filename:=rpath & "Martinrea\Tillsonburg\Vend Usage Report\Martinrea Tillsonburg Vend Usage Report " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "shelby.stewart@martinrea.com;maggie.goodeve@martinrea.com;kassandra.wilson@martinrea.com;Deidre.Ingraham@martinrea.com"
        Case 83 'Moore Packaging
            lastMonth = Format(DateAdd("m", -1, Date), "mmmm")
            'Monthly Report
            If ActiveWorkbook.name = "Moore Monthly Usage.xls" Then
                ActiveWorkbook.SaveAs Filename:=rpath & "Moore Packaging\Vend Usage Report\Monthly\Moore Monthly Usage Report - " & lastMonth & ".xlsx", FileFormat:=xlOpenXMLWorkbook
                sender = "JShand@moorepackaging.com;jneal@moorepackaging.com"
                msg = "<p>Good morning,</p> <p></p> <p>Please see attached the monthly usage for " & lastMonth & "<p></p> <p>Thank you and have a great day!</p> <p></p>" & strBuffer
            Else
                ActiveWorkbook.SaveAs Filename:=rpath & "Moore Packaging\Vend Usage Report\Moore Packaging Vend Usage Report " & lastMonth & ".xlsx", FileFormat:=xlOpenXMLWorkbook
                sender = "JShand@moorepackaging.com;jneal@moorepackaging.com"
            End If
        Case 84 'Team Oakville
            ActiveWorkbook.SaveAs Filename:=rpath & "Team Oakville\Vend Usage Report\Team Oakville Vend Usage Report " & rformat, FileFormat:=xlOpenXMLWorkbook
            sender = "Sid.Hoiting@TeamInc.com"
    End Select
    
    wbName = Left(ActiveWorkbook.name, (InStrRev(ActiveWorkbook.name, ".", -1, vbTextCompare) - 1))
    
    WatchEmails.Add thisMail
        With thisMail.TheMail
            .To = sender
            .SentOnBehalfOfName = "vending@tytanglove.ca"
            .Subject = wbName
            .BodyFormat = olFormatHTML
            .HTMLBody = msg
            .Attachments.Add ActiveWorkbook.FullName
            .Display
        End With
    
    With OutApp.ActiveWindow
        .WindowState = olNormalWindow
        .Left = 3000
        .WindowState = olMaximized
    End With
    
    Set OutApp = Nothing
    Set ExApp = Nothing
End Function
