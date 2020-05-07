Attribute VB_Name = "Module1"
Dim sIp, sPlugin, sCVSS, sRisk, sSubnet, sOc4, sIP2, sDesc, sRem, sBase1, sBase2, sBase3, sSlash1 As String
Dim iPlugStart, iPlugEnd, iFootprint, iAddDNS, iDNSEOF As Integer
Dim colPlugins As Collection

Private Sub WriteBullets()
' Commented out calls BB
'    Call SetupStrWk("StrTMP")
'    Call SetupStrWk("WkTMP")
End Sub

Sub HighlightFindings(ByVal x As Integer, ByVal y As Integer)
    For Each cell In Range("E:E")
        If Range("E" & cell.Row).Value = "" Then Exit For
        If Range("H" & cell.Row).Value = x Then
            If y = 0 Then
                Range("E" & cell.Row & ":F" & cell.Row).Select
                With Selection.Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                Range("A" & cell.Row).Value = ""
                Range("A1").Select
                Exit For
            Else
                Range("E" & cell.Row & ":F" & cell.Row).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                Range("A" & cell.Row).Value = y
                Range("A1").Select
                Exit For
            End If
        End If
    Next cell
End Sub

'Block commented out code for Strengths and Weaknesses BB
'Private Sub SetupStrWk(ByVal sSev As String)
'    Application.DisplayAlerts = False
'    If WorksheetExists(sSev) = True Then Sheets(sSev).Delete
'    Application.DisplayAlerts = True
'    Sheets.Add.Name = sSev
'    Sheets(sSev).Move After:=Sheets(5)
'    Sheets("StrengthWeakness").Select
'
'    Dim x, y, a, z As Integer
'    For Each cell In Range("C:C")
'        If sSev = "StrTMP" Then
'            If cell.Value2 = "Strengths" Then x = cell.Row
'            If cell.Value2 = "Weaknesses" Then
'                y = cell.Row
'                Exit For
'            End If
'        ElseIf sSev = "WkTMP" Then
'            If cell.Value2 = "Weaknesses" Then x = cell.Row
'            If cell.Value = "" Then
'                y = cell.Row
'                Exit For
'            End If
'        End If
'        y = y + 1
'    Next cell
'
'    z = Application.WorksheetFunction.Max(Range("A" & x & ":A" & y - 1))
'    For a = 1 To z
'        Call StrWKLoop(x, y, sSev, a)
'    Next a
'End Sub

'Private Sub StrWKLoop(ByVal x As Integer, ByVal y As Integer, ByVal sSev As String, ByVal a As Integer)
'    Dim z As String
'    For Each cell In Range("A" & x + 1 & ":A" & y - 1)
'        If cell.Value = a Then
'            z = Range("C" & cell.Row).Value
'            Sheets(sSev).Select
'            Range("A" & a).Value = a
'            Range("B" & a).Value = z
'            Sheets("StrengthWeakness").Select
'            Exit For
'        End If
'    Next cell
'End Sub

Private Sub WriteFindings()
    Application.DisplayAlerts = False
    If WorksheetExists("FindingsTMP") = True Then Sheets("FindingsTMP").Delete
    Application.DisplayAlerts = True
    Sheets.Add.Name = "FindingsTMP"
    Sheets("FindingsTMP").Move After:=Sheets(5)
    Sheets("Findings").Select
    Dim x, y, z  As Integer
    x = 1
    z = Application.WorksheetFunction.Max(Range("A:A"))
    For Each cell In Range("E:E")
        If cell.Value = "" Then Exit For
        x = x + 1
    Next cell

    For y = 1 To z
        Call WriteFindingsLoop(x, y)
    Next y
End Sub
Private Sub WriteFindingsLoop(ByVal x As Integer, ByVal y As Integer)
    Dim sProbability, sImpact, sOverall, sFinding, sRecommendation, sIPAddress As String

    For Each cell In Range("A2:A" & x + 1)
        If cell.Value = y Then
            sProbability = Range("B" & cell.Row).Value
            sImpact = Range("C" & cell.Row).Value
            sOverall = Range("D" & cell.Row).Value
            sFinding = Range("E" & cell.Row).Value
            sRecommendation = Range("F" & cell.Row).Value
            sIPAddress = Range("G" & cell.Row).Value
            
            Sheets("FindingsTMP").Select
            Range("A" & y).Value = y
            Range("B" & y).Value = sProbability
            Range("C" & y).Value = sImpact
            Range("D" & y).Value = sOverall
            Range("E" & y).Value = sFinding
            Range("F" & y).Value = sRecommendation
            Range("G" & y).Value = sIPAddress
        End If
        Sheets("Findings").Select
    Next cell
End Sub

Private Sub ParseData()
    Call CreatePluginWorksheet
    Call CreateSubnetColumn
End Sub

Private Sub SortIPAddresses()
    For i = 1 To colPlugins.Count
        Sheets(colPlugins(i)).Select
        
        Dim x, y As Integer
        For y = 2 To 256
            If Range(Replace(Split(Columns(y).Address, ":")(0), "$", "") & 1).Value = "" Then
                Exit For
            Else
                ActiveSheet.Range(Replace(Split(Columns(y).Address, ":")(0), "$", "") & ":" & Replace(Split(Columns(y).Address, ":")(0), "$", "")).RemoveDuplicates Columns:=1, Header:=xlYes
                Columns(Replace(Split(Columns(y).Address, ":")(0), "$", "") & ":" & Replace(Split(Columns(y).Address, ":")(0), "$", "")).Select
                
                ActiveWorkbook.Worksheets(colPlugins(i)).Sort.SortFields.Clear
                ActiveWorkbook.Worksheets(colPlugins(i)).Sort.SortFields.Add Key:=Range(Replace(Split(Columns(y).Address, ":")(0), "$", "") & "1"), _
                    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                With ActiveWorkbook.Worksheets(colPlugins(i)).Sort
                    .SetRange Range(Replace(Split(Columns(y).Address, ":")(0), "$", "") & "2:" & Replace(Split(Columns(y).Address, ":")(0), "$", "") & "257")
                    .Header = xlGuess
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
            End If
        Next y
    Next
End Sub

Private Sub PopulateTable()
    'Loop worksheets
    For i = 1 To colPlugins.Count
        Sheets(colPlugins(i)).Select
        Dim x, y As Integer
        
        'Loop Columns
        Dim colIP As Collection
        Set colIP = New Collection
        For y = 2 To 256
            If Range(Replace(Split(Columns(y).Address, ":")(0), "$", "") & 1).Value = "" Then
                Exit For
            Else
                sIP2 = Range(Replace(Split(Columns(y).Address, ":")(0), "$", "") & 1).Value & "." & Range(Replace(Split(Columns(y).Address, ":")(0), "$", "") & 2).Value
                
                'Loop Rows
                For x = 3 To 257
                    If Range(Replace(Split(Columns(y).Address, ":")(0), "$", "") & x).Value = "" Then
                        Exit For
                    Else
                        sIP2 = sIP2 & ", " & Range(Replace(Split(Columns(y).Address, ":")(0), "$", "") & x).Value
                    End If
                Next x
            End If
            colIP.Add sIP2
        Next y
        
        sPlugin = colPlugins(i)
        sDesc = ""
        sRem = ""
        sRisk = Range("A2").Value
        sCVSS = Range("A1").Value
        
        Sheets("Vuln").Select
        For Each cell In Range("A:A")
            If cell.Value = "" Then Exit For
            If cell.Value = CStr(colPlugins(i)) Then
                'sPlugin = Range("A" & cell.Row).Value
                sDesc = Range("B" & cell.Row).Value
                sRem = Range("C" & cell.Row).Value
                'sRisk = Range("D" & cell.Row).Value
                'sCVSS = Range("E" & cell.Row).Value
                Exit For
            End If
        Next cell
        
        Sheets("VulnReported").Select
        Call PlugData(colIP)
    Next
End Sub

Private Sub Cleanup()
    For i = 1 To colPlugins.Count
        Application.DisplayAlerts = False
        Sheets(colPlugins(i)).Delete
        Application.DisplayAlerts = True
    Next i
End Sub

Private Sub PlugData(ByVal colIP As Collection)
    Dim x As Integer
    x = 1
    For Each cell In Range("A:A")
        If cell.Value = "" Then Exit For
        x = x + 1
    Next cell

    Range("A" & x).Value = sPlugin
    Range("B" & x).Value = sDesc
    Range("C" & x).Value = sRem
    Range("D" & x).Value = sRisk & "/" & sCVSS
    
    For i = 1 To colIP.Count
        Dim z As String
        If colIP.Count = 1 Then
            z = colIP(i)
        Else
            If i = 1 Then
                z = colIP(i)
            Else
                z = z & vbCrLf & colIP(i)
            End If
        End If
    Next i
    Range("E" & x).Value = z
    Range("F" & x).Value = sCVSS
End Sub

Private Sub WriteObservations()
Application.DisplayAlerts = True
Set colPlugins = New Collection
Sheets("Nessus").Select
For Each cell In Range("A:A")
    If cell.Value = "</Report>" Then Exit For
    If InStr(1, cell.Value, "host-ip", vbTextCompare) > 0 Then
        sIp = Mid(cell.Value, 21, Len(cell.Value) - 26)
    End If
    If (InStr(1, cell.Value, "ReportItem", vbTextCompare) > 0) And (InStr(1, cell.Value, "pluginID", vbTextCompare) > 0) Then
        iPlugStart = InStr(1, cell.Value, "pluginID", vbTextCompare) + 10
        iPlugEnd = InStr(iPlugStart + 1, cell.Value, """", vbTextCompare)
        sPlugin = Mid(cell.Value, iPlugStart, iPlugEnd - iPlugStart)
    End If
    'getting the scvss at this point may not be necessary
    If InStr(1, cell.Value, "cvss_base_score", vbTextCompare) > 0 Then
        sCVSS = Mid(cell.Value, 18, Len(cell.Value) - 35)
    End If
    If InStr(1, cell.Value, "risk_factor", vbTextCompare) > 0 Then
        sRisk = Mid(cell.Value, 14, Len(cell.Value) - 27)
        If sRisk = "Critical" Then sRisk = "High"
        
        If sRisk <> "None" Then
            'Determine subnet and fourth octet
            Dim x As Integer
            x = Len(sIp)
            Do
                If Mid(sIp, x, 1) = "." Then
                    sSubnet = Left(sIp, x - 1)
                    sOc4 = Right(sIp, Len(sIp) - x)
                    Exit Do
                End If
            x = x - 1
            Loop Until x = 1

            'Begin sorting through data
            Call ParseData
        End If
    End If
Next cell

Call SortIPAddresses
Call PopulateTable
Call Cleanup
Call OrderData
End Sub

Private Sub OrderData()
    Sheets("VulnReported").Select
    Dim x As Integer
    x = 1
    For Each cell In Range("D:D")
        If cell.Value = "" Then Exit For
        x = x + 1
    Next cell
    
    ActiveWorkbook.Worksheets("VulnReported").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("VulnReported").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("F1:F" & x), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("VulnReported").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A2").Select
End Sub

Private Sub CreateSubnetColumn()
    Sheets(sPlugin).Select
    Dim y As Integer
    Dim z As String
    z = ""
    Range("A:A").Select
    Selection.NumberFormat = "@"
    Range("A1").Value = sCVSS
    Range("A2").Value = sRisk
    For y = 2 To 256
        If Range(Replace(Split(Columns(y).Address, ":")(0), "$", "") & 1).Value = "" Then Exit For
    Next y
    
    If y = 2 Then
        'Use first column if worksheet is blank
        Range(Replace(Split(Columns(y).Address, ":")(0), "$", "") & 1).Value = sSubnet
        Call ListIP(y)
    Else
        'Look for existing subnet
        Dim x As Integer
        For x = 1 To y
            If Range(Replace(Split(Columns(x).Address, ":")(0), "$", "") & 1).Value = sSubnet Then
                Call ListIP(x)
                z = sSubnet
                Exit For
            End If
        Next x
        
        'Use new column if no existing subnet found
        If z = "" Then
            Range(Replace(Split(Columns(y).Address, ":")(0), "$", "") & 1).Value = sSubnet
            Call ListIP(y)
        End If
    End If
End Sub

Private Function ListIP(ByVal y As Integer)
    Dim x As Integer
    Dim z As String
    x = 2
    z = ""
    
    If Range(Replace(Split(Columns(y).Address, ":")(0), "$", "") & 2).Value = "" Then
        Range(Replace(Split(Columns(y).Address, ":")(0), "$", "") & 2).Value = sOc4
    Else
        For Each cell In Range(Replace(Split(Columns(y).Address, ":")(0), "$", "") & 2 & ":" & Replace(Split(Columns(y).Address, ":")(0), "$", "") & 257)
            If Range(Replace(Split(Columns(y).Address, ":")(0), "$", "") & cell.Row).Value = "" Then
                Exit For
            ElseIf Range(Replace(Split(Columns(y).Address, ":")(0), "$", "") & cell.Row).Value = sOc4 Then
                z = sOc4
                Exit For
            Else
                 x = x + 1
            End If
        Next cell
        If z = "" Then
            Range(Replace(Split(Columns(y).Address, ":")(0), "$", "") & x).Value = sOc4
        End If
    End If
End Function

Private Sub CreatePluginWorksheet()
    If Worksheets.Count = 3 Then
        Sheets.Add.Name = sPlugin
        colPlugins.Add sPlugin
    Else
        Dim x As Integer
        Dim y As String
        y = ""
        For x = 1 To Worksheets.Count
            If Worksheets(x).Name = sPlugin Then
                y = Worksheets(x).Name
            End If
        Next x
        If y = "" Then
            Sheets.Add.Name = sPlugin
            colPlugins.Add sPlugin
        End If
    End If
End Sub

Private Sub CopyFiles(ByVal sEng As String, ByVal sEng2 As String)
    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    sBase1 = Range("D1").Value
    If Mid(sBase1, Len(sBase1), 1) <> "\" Then sBase1 = sBase1 & "\"
    sBase2 = Range("D2").Value
    sBase3 = Range("D3").Value
    
    Dim boolCopy As Boolean
    boolCopy = False
    If boolCopy = False Then
        If sEng2 = "" Then
            If InStr(1, sBase1, "\", vbTextCompare) Then
                sSlash1 = "\"
            Else
                sSlash1 = "/"
            End If
            Call fso.CopyFile(sBase1 & sBase2 & sSlash1 & "RM0575-IT-FI-" & sEng & ".docx", sBase1 & sBase3 & sSlash1 & "RM0575-IT-FI-" & sEng & ".docx")
        Else
            If InStr(1, sBase1, "\", vbTextCompare) Then
                sSlash1 = "\"
            Else
                sSlash1 = "/"
            End If
            Call fso.CopyFile(sBase1 & sBase2 & sSlash1 & sEng, sEng2)
        End If
    Else
        If sEng2 = "" Then
            If InStr(1, sBase1, "\", vbTextCompare) Then
                sSlash1 = "\"
            Else
                sSlash1 = "/"
            End If
            FileCopy sBase1 & sBase2 & sSlash1 & "RM0575-IT-FI-" & sEng & ".docx", sBase1 & sBase3 & sSlash1 & "RM0575-IT-FI-" & sEng & ".docx"
        Else
            If InStr(1, sBase1, "\", vbTextCompare) Then
                sSlash1 = "\"
            Else
                sSlash1 = "/"
            End If
            Call fso.CopyFile(sBase1 & sBase2 & sSlash1 & sEng, sEng2)
        End If
    End If
End Sub

Private Sub DeleteTempSheets(ByVal sTempVulnSheet As String)
    Application.DisplayAlerts = False
    On Error Resume Next
    If WorksheetExists("StrTMP") = True Then Sheets("StrTMP").Delete
    If WorksheetExists("WkTMP") = True Then Sheets("WkTMP").Delete
    If WorksheetExists("FindingsTMP") = True Then Sheets("FindingsTMP").Delete
    If WorksheetExists("DomainTMP") = True Then Sheets("DomainTMP").Delete
    If WorksheetExists(sTempVulnSheet) = True Then Sheets(sTempVulnSheet).Delete
    Application.DisplayAlerts = True
End Sub

Private Sub ReformatNOWP(ByRef wdoc, ByRef sEng, ByRef WordContent)
    'Remove unnecessary formatting in cover letters
    wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
    For x = 1 To 2
        wdoc.ActiveWindow.Selection.Find.Text = "[consultant name]"
        wdoc.ActiveWindow.Selection.Find.Execute
        wdoc.ActiveWindow.Selection.Font.Italic = False
        wdoc.ActiveWindow.Selection.Font.Underline = wdUnderlineNone
        wdoc.ActiveWindow.Selection.MoveRight Unit:=wdCharacter, Count:=1

        wdoc.ActiveWindow.Selection.Find.Text = "[phone number]"
        wdoc.ActiveWindow.Selection.Find.Execute
        wdoc.ActiveWindow.Selection.Font.Italic = False
        wdoc.ActiveWindow.Selection.Font.Underline = wdUnderlineNone
        wdoc.ActiveWindow.Selection.MoveRight Unit:=wdCharacter, Count:=1
        
        wdoc.ActiveWindow.Selection.Find.Text = "[email address]"
        wdoc.ActiveWindow.Selection.Find.Execute
        wdoc.ActiveWindow.Selection.Font.Italic = False
        wdoc.ActiveWindow.Selection.Font.Underline = wdUnderlineNone
        wdoc.ActiveWindow.Selection.MoveRight Unit:=wdCharacter, Count:=1
    Next x
    
    wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
    Select Case sEng
        Case "SE"
            'remove statement that is likely never going to be used
            wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
            wdoc.ActiveWindow.Selection.Find.Text = "This engagement is a dual-purpose engagement, covering both the FFIEC regulatory procedures and it will be used by the financial auditors of BKD in the performance of the external financial statement audit.(A)"
            wdoc.ActiveWindow.Selection.Find.Execute
            wdoc.ActiveWindow.Selection.TypeBackspace
            wdoc.ActiveWindow.Selection.TypeBackspace
            wdoc.ActiveWindow.Selection.TypeBackspace
            
            wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
            wdoc.ActiveWindow.Selection.Find.Text = "(A)"
            wdoc.ActiveWindow.Selection.Find.Execute
            wdoc.ActiveWindow.Selection.MoveLeft Unit:=wdCharacter, Count:=1
            wdoc.ActiveWindow.Selection.MoveDown Unit:=wdLine, Count:=7, Extend:=wdExtend
            wdoc.ActiveWindow.Selection.TypeBackspace
            With WordContent.Find
                .Text = "(a)"
                .Replacement.Text = ""
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
            
            'highlight section that may be missed
            wdoc.ActiveWindow.Selection.Find.Text = "[A list of 'cracked' domain credentials is included in Section VI. OR No domain credentials were captured as noted in Section VI.]"
            wdoc.ActiveWindow.Selection.Find.Execute
            wdoc.ActiveWindow.Selection.MoveLeft Unit:=wdCharacter, Count:=1
            wdoc.ActiveWindow.Selection.Delete
            wdoc.ActiveWindow.Selection.Find.Text = "A list of 'cracked' domain credentials is included in Section VI. OR No domain credentials were captured as noted in Section VI."
            wdoc.ActiveWindow.Selection.Find.Execute
            wdoc.ActiveWindow.Selection.Font.Italic = False
            wdoc.ActiveWindow.Selection.Font.Underline = wdUnderlineNone
            wdoc.ActiveWindow.Selection.Range.HighlightColorIndex = wdYellow
            wdoc.ActiveWindow.Selection.MoveRight Unit:=wdCharacter, Count:=1
            wdoc.ActiveWindow.Selection.Delete
            wdoc.ActiveWindow.Selection.MoveRight Unit:=wdCharacter, Count:=1
            wdoc.ActiveWindow.Selection.MoveDown Unit:=wdLine, Count:=1
            wdoc.ActiveWindow.Selection.MoveDown Unit:=wdLine, Count:=15, Extend:=wdExtend
            wdoc.ActiveWindow.Selection.Font.Italic = False
            wdoc.ActiveWindow.Selection.Font.Underline = wdUnderlineNone
            wdoc.ActiveWindow.Selection.Range.HighlightColorIndex = wdYellow
            
            'Remove unnecessary "insert management response here" note
            wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
            wdoc.ActiveWindow.Selection.Find.Text = "[Insert management response]"
            wdoc.ActiveWindow.Selection.Find.Execute
            wdoc.ActiveWindow.Selection.TypeBackspace
        Case Else
    End Select
End Sub

Private Sub ImportBKDMan(ByRef wdoc, ByRef sEng)
    'Import BKDMan text from the textbox on the Control dashboard of the Excel file
    wdoc.ActiveWindow.Selection.Find.Text = "[Insert report from BKDMAN Section 3749.302 for financial institutions]"
    wdoc.ActiveWindow.Selection.Find.Execute
    wdoc.ActiveWindow.Selection.TypeBackspace
    wdoc.ActiveWindow.Selection.TypeText Text:=Worksheets("Control").txtBKDMan.Text
    wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
    
    'Move to Report Letter
    wdoc.ActiveWindow.Selection.Find.Text = "Report Letter"
    wdoc.ActiveWindow.Selection.Find.Execute
    wdoc.ActiveWindow.Selection.Find.Execute
    wdoc.ActiveWindow.Selection.MoveRight
    
    'Modify based on engagement type
    Select Case sEng
        Case "SE"
            wdoc.ActiveWindow.Selection.Find.Text = "[Insert Section]"
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
                wdoc.ActiveWindow.Selection.TypeText Text:="Section II"
            wdoc.ActiveWindow.Selection.Find.Text = "[Client Name]"
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
                wdoc.ActiveWindow.Selection.TypeText Text:=Range("A2").Value
            wdoc.ActiveWindow.Selection.Find.Text = "(the Bank)"
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
                wdoc.ActiveWindow.Selection.TypeText Text:="(the " & Range("A7").Value & ")"
            wdoc.ActiveWindow.Selection.Find.Text = "[EL date]"
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
                wdoc.ActiveWindow.Selection.TypeText Text:=" " & Range("A4").Value
            wdoc.ActiveWindow.Selection.Find.Text = "[specify – internal control over information technology systems OR internal network security OR social engineering awareness]"
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
                wdoc.ActiveWindow.Selection.TypeText Text:="social engineering awareness"
            wdoc.ActiveWindow.Selection.Find.Text = "[date = last day of fieldwork]"
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
                wdoc.ActiveWindow.Selection.TypeText Text:=" " & Range("A6").Value
            wdoc.ActiveWindow.Selection.Find.Text = "[Insert Section]"
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
                wdoc.ActiveWindow.Selection.TypeText Text:="Section II"
            wdoc.ActiveWindow.Selection.Find.Text = "[Select one of the following sentences – "
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
            wdoc.ActiveWindow.Selection.Find.Text = "(specify - findings, recommendations and observations)"
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
                wdoc.ActiveWindow.Selection.TypeText Text:=" recommendations"
            wdoc.ActiveWindow.Selection.Find.Text = "[Insert Section]"
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
                wdoc.ActiveWindow.Selection.TypeText Text:=" Section II"
            wdoc.ActiveWindow.Selection.Find.Text = " OR There were no findings identified in connection with the procedures performed.]"
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
        Case Else
    End Select
    
    'Remove blank line
    wdoc.ActiveWindow.Selection.Find.Text = "________________, ______"
        wdoc.ActiveWindow.Selection.Find.Execute
        wdoc.ActiveWindow.Selection.Delete
        wdoc.ActiveWindow.Selection.Delete
        wdoc.ActiveWindow.Selection.Delete
    wdoc.ActiveWindow.Selection.Find.Text = "[Date]"
        wdoc.ActiveWindow.Selection.Find.Execute
        wdoc.ActiveWindow.Selection.TypeBackspace
        wdoc.ActiveWindow.Selection.TypeText Text:=Range("A16").Text
End Sub


Private Sub WriteReport(ByVal sEng As String)
    Dim sReportLocation As String
    sReportLocation = FileSelectBox("*.docx")
    Call DeleteTempSheets("TMP")
    If sEng = "IPT" Then Call WriteFindings
    Sheets("Control").Select
    Dim observation, recommendation, ipaddress, cvss, FilePath, sType, sClient, sAddress, sElDate, sDateFirst, sDateLast, sTempVulnSheet As String
    Dim sProbability, sImpact, sOverall, sFinding, sRecommendation, sIPAddress, sStrBullet, sWkBullet As String
    Dim x, y, z As Integer
    Dim currentRange As Range
    'filepath = sBase1 & sBase3 & sSlash1 & "RM0575-IT-FI-" & sEng & ".docx"
    FilePath = sReportLocation
    'Call CopyFiles(sEng, "")
    
    Dim wdApp As Object
    Dim wdoc As Object
    
    Set wdApp = CreateObject("word.application")
    wdApp.Visible = True
    'Set wdoc = wdApp.Documents.Open(sBase1 & sBase3 & sSlash1 & "RM0575-IT-FI-" & sEng & ".docx")
    Set wdoc = wdApp.Documents.Open(sReportLocation)
    Set WordContent = wdoc.Content
    wdoc.TrackRevisions = True

    'Import BKDMan
    Call ImportBKDMan(wdoc, sEng)
    
    'Reformat optional sections
    Call ReformatNOWP(wdoc, sEng, WordContent)
    
    'Write headers and footers
    Dim oSection As Word.Section
    Dim oRange As Word.Range
    For Each oSection In wdoc.Sections()
        For y = 1 To 3
            Set oRange = oSection.Footers(y).Range
            oRange.Find.Execute FindText:="{client_name}", _
            ReplaceWith:=Range("A2").Value, Replace:=wdReplaceAll
            Set oRange = Nothing
        Next
' BB
'        For y = 1 To 3
'            Set oRange = oSection.Footers(y).Range
'            oRange.Find.Execute FindText:="{date_last}", _
'            ReplaceWith:=Range("A6").Value, Replace:=wdReplaceAll
'            Set oRange = Nothing
'        Next
    Next
    
    'Write report
    For Each cell In Range("B2:B32767")
        If cell.Value = "" Then Exit For
        With WordContent.Find
            .Text = Range("B" & cell.Row).Value
            
            If Range("B" & cell.Row).Value = "{today}" Then
                .Replacement.Text = Format(Range("A" & cell.Row).Value, "MMMM DD, YYYY")
            ElseIf Range("B" & cell.Row).Value = "{date procedures are substantially completed}" Then
                .Replacement.Text = Format(Range("A" & cell.Row).Value, "MMMM DD, YYYY")
            ElseIf Range("B" & cell.Row).Value = "{tech_contact}" Then
                If Range("A" & cell.Row).Value = "" Then
                    .Replacement.Text = Range("A8").Value
                Else
                    .Replacement.Text = Range("A" & cell.Row).Value
                End If
            ElseIf Range("B" & cell.Row).Value = "{tech_title}" Then
                If Range("A" & cell.Row).Value = "" And Range("A" & cell.Row - 1).Value = "" Then
                     If Range("A9").Value <> "" Then
                        .Replacement.Text = ", " & Range("A9").Value & ","
                    Else
                        .Replacement.Text = ""
                    End If
                Else
                    If Range("A" & cell.Row).Value <> "" Then
                        .Replacement.Text = ", " & Range("A" & cell.Row).Value
                    Else
                        .Replacement.Text = ""
                    End If
                End If
            Else
                .Replacement.Text = Range("A" & cell.Row).Value
            End If
            
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    Next cell
    wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
    
    'Write observations for pen tests only
    If sEng = "EPT" Or sEng = "EVS" Or sEng = "IPT" Then
        Sheets("VulnReported").Select
        x = 2
        For Each cell In Range("A:A")
            If cell.Value = "" Then Exit For
            observation = Range("b" & x).Value
            recommendation = Range("c" & x).Value
            ipaddress = Range("e" & x).Value
            cvss = Range("d" & x).Value
            If cvss = "" Then Exit For
            
            wdoc.ActiveWindow.Selection.Find.Text = "{cvss" & x - 1 & "}"
            wdoc.ActiveWindow.Selection.Find.Execute
            wdoc.ActiveWindow.Selection.TypeBackspace
            wdoc.ActiveWindow.Selection.TypeText Text:=cvss
            wdoc.ActiveWindow.Selection.Find.Text = "{observation" & x - 1 & "}"
            wdoc.ActiveWindow.Selection.Find.Execute
            wdoc.ActiveWindow.Selection.TypeBackspace
            wdoc.ActiveWindow.Selection.TypeText Text:=observation
            wdoc.ActiveWindow.Selection.Find.Text = "{ipaddress" & x - 1 & "}"
            wdoc.ActiveWindow.Selection.Find.Execute
            wdoc.ActiveWindow.Selection.TypeBackspace
            wdoc.ActiveWindow.Selection.TypeText Text:=ipaddress
            wdoc.ActiveWindow.Selection.Find.Text = "{recommendation" & x - 1 & "}"
            wdoc.ActiveWindow.Selection.Find.Execute
            wdoc.ActiveWindow.Selection.TypeBackspace
            wdoc.ActiveWindow.Selection.TypeText Text:=recommendation
        
            x = x + 1
        Next cell
        wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
    End If
    'Write External Pen Test Footprinting
    If sEng = "EPT" Then
        Dim wTbl As Object
        
        If WorksheetExists("Footprinting") = True Then
            Sheets("Footprinting").Select
            For Each cell In Range("A:A")
                If cell.Value = "" And Range("A" & cell.Row + 1).Value = "" Then
                    x = cell.Row - 1
                    Exit For
                End If
            Next cell
            If x > 0 Then
                Range("A1:B" & x).Copy
                wdoc.ActiveWindow.Selection.Find.Text = "{footprinting}"
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
                wdoc.ActiveWindow.Selection.Paste
                'wdoc.ActiveWindow.Selection.TypeText Text:=cvss
            End If
        End If
    
    End If
    'Write findings for IPT only
    If sEng = "IPT" Then
        Call WriteBullets
        Sheets("FindingsTMP").Select
        x = 1
        z = Application.WorksheetFunction.Max(Range("A:A"))
        If z <> 0 Then
            For Each cell In Range("A" & x & ":A" & z)
                If cell.Value = "" Then Exit For
                sProbability = Range("B" & x).Value
                sImpact = Range("C" & x).Value
                sOverall = Range("D" & x).Value
                sFinding = Range("E" & x).Value
                sRecommendation = Range("F" & x).Value
                sIPAddress = Range("G" & x).Value
                
                sProbability = Trim(sProbability & vbNullString)
                sImpact = Trim(sImpact & vbNullString)
                sOverall = Trim(sOverall & vbNullString)
                sFinding = Trim(sFinding & vbNullString)
                sRecommendation = Trim(sRecommendation & vbNullString)
                sIPAddress = Trim(sIPAddress & vbNullString)
'changed Risk Rating to just risk to match new report BB
                wdoc.ActiveWindow.Selection.Find.Text = "{risk" & x & "}"
                wdoc.ActiveWindow.Selection.Find.Execute
'commented out the next line BB
'                wdoc.ActiveWindow.Selection.TypeText Text:="Risk Rating"
                wdoc.ActiveWindow.Selection.EndKey Unit:=wdLine
                If sOverall = "Medium" Then
'Changed Text:=Chr(117) to Text:=Medium, High, Critical etc
                    wdoc.ActiveWindow.Selection.TypeText Text:="Medium"
                ElseIf sOverall = "High" Then
                    wdoc.ActiveWindow.Selection.TypeText Text:="High"
'                    wdoc.ActiveWindow.Selection.TypeText Text:=Chr(117)
                ElseIf sOverall = "Critical" Then
                    wdoc.ActiveWindow.Selection.TypeText Text:="Critical"
'                    wdoc.ActiveWindow.Selection.TypeText Text:=Chr(117)
'                    wdoc.ActiveWindow.Selection.TypeText Text:=Chr(117)
                End If
                
                wdoc.ActiveWindow.Selection.Find.Text = "{probability" & x & "}"
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
                wdoc.ActiveWindow.Selection.TypeText Text:=sProbability
                wdoc.ActiveWindow.Selection.Find.Text = "{impact" & x & "}"
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
                wdoc.ActiveWindow.Selection.TypeText Text:=sImpact
                wdoc.ActiveWindow.Selection.Find.Text = "{finding" & x & "}"
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
                wdoc.ActiveWindow.Selection.TypeText Text:=sFinding
'Changed "finding_ip" to "ipaddress" to match new report BB
                wdoc.ActiveWindow.Selection.Find.Text = "{ipaddress" & x & "}"
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
                wdoc.ActiveWindow.Selection.TypeText Text:=sIPAddress
'Changed "finding_recommendation" to "recommendation" to match new report BB
                wdoc.ActiveWindow.Selection.Find.Text = "{recommendation" & x & "}"
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
                wdoc.ActiveWindow.Selection.TypeText Text:=sRecommendation
            
                x = x + 1
                wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
            Next cell
        End If
        
        'Write bullets for IPT only
' Block commented out code for Strengths and Weaknesses BB
'        Sheets("StrTMP").Select
'        x = 1
'        z = Application.WorksheetFunction.Max(Range("A:A"))
'        If z <> 0 Then
'            For Each cell In Range("A1:A" & z)
'                'If cell.Value = "" Then Exit For
'                sStrBullet = Range("B" & x).Value
'
'                wdoc.ActiveWindow.Selection.Find.Text = "{strength" & x & "}"
'                wdoc.ActiveWindow.Selection.Find.Execute
'                wdoc.ActiveWindow.Selection.TypeBackspace
'                wdoc.ActiveWindow.Selection.TypeText Text:=sStrBullet
'
'                x = x + 1
'            Next cell
'        End If
'        wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
'
'        Sheets("WkTMP").Select
'        x = 1
'        z = Application.WorksheetFunction.Max(Range("A:A"))
'        If z <> 0 Then
'            For Each cell In Range("A1:A" & z)
'                'If cell.Value = "" Then Exit For
'                sWkBullet = Range("B" & x).Value
'
'                wdoc.ActiveWindow.Selection.Find.Text = "{weakness" & x & "}"
'                wdoc.ActiveWindow.Selection.Find.Execute
'                wdoc.ActiveWindow.Selection.TypeBackspace
'                wdoc.ActiveWindow.Selection.TypeText Text:=sWkBullet
'
'                x = x + 1
'            Next cell
'        End If
        wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
        
        Sheets("Control").Select
        wdoc.ActiveWindow.Selection.Find.Text = "{summary}"
        wdoc.ActiveWindow.Selection.Find.Execute
        wdoc.ActiveWindow.Selection.TypeBackspace
        wdoc.ActiveWindow.Selection.TypeText Text:=Worksheets("Control").txtSummary.Text
        wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
        
        Sheets("Control").Select
        Dim sFilename As String
        sFilename = UCase(Left(Range("A2").Value, 1))
        Dim i As Integer
        For i = 1 To Len(Range("A2").Value)
            If Mid(Range("A2").Value, i, 1) = " " Then
                sFilename = sFilename & UCase(Mid(Range("A2").Value, i + 1, 1))
            End If
        Next i
        sFilename = sFilename & "_Vulnerabilities"
        Sheets("VulnReported").Select
        Sheets("VulnReported").Copy Before:=Sheets(1)
        If sEng = "IPT" Then
            ActiveSheet.Name = "Internal Penetration Test"
        ElseIf sEng = "IVS" Then
            ActiveSheet.Name = "Internal Vulnerability Scan"
        ElseIf sEng = "EPT" Then
            ActiveSheet.Name = "External Penetration Test"
        ElseIf sEng = "EVS" Then
            ActiveSheet.Name = "External Vulnerability Scan"
        End If
        sTempVulnSheet = ActiveSheet.Name
        Columns("A:A").Select
        Selection.Delete Shift:=xlToLeft
        Columns("E:E").Select
        Selection.Delete Shift:=xlToLeft
        Range("A1").Select
        Sheets(ActiveSheet.Name).Select
        x = 0
        For Each cell In Range("E:E")
            If cell.Value = "" Then Exit For
            x = x + 1
        Next cell
        
        Call SaveAs(ActiveSheet.Name, ThisWorkbook.Path, sFilename)
        Call DeleteTempSheets(sTempVulnSheet)
    End If
    
    'Write Social Engineering
    If sEng = "SE" Then
        Sheets("Credentials").Select
        x = 0
        For Each cell In Range("A:A")
            If cell.Value = "" Then Exit For
            x = x + 1
        Next cell
        Sheets("Control").Select
        Range("H18").Value = x
        
        x = 1
        For Each cell In Range("H7:H8")
            'If cell.Value = "" Then Exit For
            wdoc.ActiveWindow.Selection.Find.Text = "{screenshot" & x & "}"
            wdoc.ActiveWindow.Selection.Find.Execute
            wdoc.ActiveWindow.Selection.TypeBackspace
            If cell.Value <> "" Then wdoc.ActiveWindow.Selection.InlineShapes.AddPicture FileName:= _
                Range("H" & cell.Row).Value, LinkToFile:=False, _
                SaveWithDocument:=True
        
            x = x + 1
        Next cell
        wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
        
        For Each cell In Range("G9:G32767")
            If cell.Value = "" Then Exit For
            If Range("G" & cell.Row).Value = "{email_body}" Then
                wdoc.ActiveWindow.Selection.Find.Text = "{email_body}"
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
                wdoc.ActiveWindow.Selection.TypeText Text:=Range("H" & cell.Row).Value
            ElseIf Range("G" & cell.Row).Value = "{percent}" Then
                Dim sDec As String
                sDec = Round(Range("H" & cell.Row).Value * 100, 2)
                If CDec(sDec) < 1 And CDec(sDec) > 0 And Len(sDec) = 3 Then
                    sDec = "0.00%"
                ElseIf CDec(sDec) = 100 Then
                    sDec = "100%"
                Else
                    sDec = sDec & "%"
                End If
                wdoc.ActiveWindow.Selection.Find.Text = "{percent}"
                wdoc.ActiveWindow.Selection.Find.Execute
                wdoc.ActiveWindow.Selection.TypeBackspace
                wdoc.ActiveWindow.Selection.TypeText Text:=" " & sDec
            ElseIf Range("G" & cell.Row).Value = "{clicks}" Or Range("G" & cell.Row).Value = "{submits}" Or Range("G" & cell.Row).Value = "{valid}" Or Range("G" & cell.Row).Value = "{total_users}" Then
                With WordContent.Find
                    .Text = Range("G" & cell.Row).Value
                    Select Case Range("H" & cell.Row).Value
                        Case 0
                            .Replacement.Text = "zero"
                        Case 1
                            .Replacement.Text = "one"
                        Case 2
                            .Replacement.Text = "two"
                        Case 3
                            .Replacement.Text = "three"
                        Case 4
                            .Replacement.Text = "four"
                        Case 5
                            .Replacement.Text = "five"
                        Case 6
                            .Replacement.Text = "six"
                        Case 7
                            .Replacement.Text = "seven"
                        Case 8
                            .Replacement.Text = "eight"
                        Case 9
                            .Replacement.Text = "nine"
                        Case Else
                            .Replacement.Text = Range("H" & cell.Row).Value
                    End Select
                    
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                End With
            Else
                With WordContent.Find
                    .Text = Range("G" & cell.Row).Value
                    .Replacement.Text = Range("H" & cell.Row).Value
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                End With
            End If
        wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
        Next cell
        
        Sheets("Emails").Select
        wdoc.ActiveWindow.Selection.Find.Text = "{email_list}"
        wdoc.ActiveWindow.Selection.Find.Execute
        wdoc.ActiveWindow.Selection.TypeBackspace
        For Each cell In Range("A:A")
            If cell.Value = "" Then Exit For
            wdoc.ActiveWindow.Selection.TypeText Text:=cell.Value
            If Range("A" & cell.Row + 1).Value <> "" Then wdoc.ActiveWindow.Selection.TypeText Text:=Chr(10)
        Next cell
        
        Sheets("Credentials").Select
        For Each cell In Range("D:D")
            If cell.Value = "" Then Exit For
            z = z + 1
        Next cell
        
        If z > 0 Then
            Columns("A:A").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            For Each cell In Range("A1:A" & z)
                Range("A" & cell.Row).Value = cell.Row
            Next cell
            Columns("E:E").Select
            
            ActiveWorkbook.Worksheets("Credentials").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Credentials").Sort.SortFields.Add Key:=Range("E1") _
                , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            With ActiveWorkbook.Worksheets("Credentials").Sort
                .SetRange Range("A1:G" & z)
                .Header = xlNo
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            
            wdoc.ActiveWindow.Selection.Find.Text = "{usernames}"
            wdoc.ActiveWindow.Selection.Find.Execute
            wdoc.ActiveWindow.Selection.TypeBackspace
            For Each cell In Range("E:E")
                If cell.Value = "" Then Exit For
                wdoc.ActiveWindow.Selection.TypeText Text:=cell.Value
                If cell.Row <> z Then wdoc.ActiveWindow.Selection.TypeText Text:=Chr(10)
            Next cell
            
            ActiveWorkbook.Worksheets("Credentials").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Credentials").Sort.SortFields.Add Key:=Range("A1") _
                , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            With ActiveWorkbook.Worksheets("Credentials").Sort
                .SetRange Range("A1:G" & z)
                .Header = xlNo
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            
            Columns("A:A").Select
            Selection.Delete Shift:=xlToLeft
            Range("A1").Select
            Sheets("Control").Select
        Else
            wdoc.ActiveWindow.Selection.Find.Text = "{usernames}"
            wdoc.ActiveWindow.Selection.Find.Execute
            wdoc.ActiveWindow.Selection.TypeBackspace
            wdoc.ActiveWindow.Selection.TypeText Text:="No usernames were captured during testing."
        End If
    End If

    MsgBox "Please take a moment to clean up the report before pressing 'OK.'" & vbCrLf & vbCrLf & "Once all corrections have been made, this macro will complete its changes, including updating all page numbers.", , "Update report"
    wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
    Call SetPageNumbers(wdoc)
    Sheets("Control").Select
    Call CustomMailMessage(sEng)
End Sub

Private Sub SaveAs(ByVal sSheetName As String, ByVal sPath As String, ByVal sFilename As String)
    Dim NewBook As Workbook
    Set NewBook = Workbooks.Add
    ThisWorkbook.Sheets(sSheetName).Copy Before:=NewBook.Sheets(1)
    Dim i As Integer
    For i = Sheets.Count To 1 Step -1
        Application.DisplayAlerts = False
        If Sheets(i).Name <> sSheetName Then Sheets(i).Delete
        Application.DisplayAlerts = True
    Next i
    
    Application.DisplayAlerts = False
    NewBook.SaveAs FileName:=sPath & "\" & sFilename
    NewBook.Close (True)
    Application.DisplayAlerts = True
End Sub

Private Sub LoopEmails(ByVal x As Integer, ByVal z As String, ByRef wdoc As Object, ByVal b As Boolean)
    Dim y As Integer
    y = x
    wdoc.ActiveWindow.Selection.Find.Text = z
    wdoc.ActiveWindow.Selection.Find.Execute
    wdoc.ActiveWindow.Selection.TypeBackspace
    For Each cell In Range("A1:A" & x)
        If cell.Value = "" Then Exit For
        wdoc.ActiveWindow.Selection.TypeText Text:=Range("A" & cell.Row).Value
        If Range("A" & cell.Row + 1).Value <> "" Then
            If cell.Row <> 42 Then
                If y <> 1 Then wdoc.ActiveWindow.Selection.TypeText Text:=Chr(10)
                y = y - 1
            Else
                If Range("A85").Value <> "" Then
                    wdoc.ActiveWindow.Selection.TypeText Text:=Chr(10)
                    If b = False Then
                        wdoc.ActiveWindow.Selection.TypeText Text:="{email_list}"
                    Else
                        wdoc.ActiveWindow.Selection.TypeText Text:="{email_list2}"
                    End If
                End If
            End If
        End If
    Next cell
    Rows("1:" & x).Select
    Selection.Delete Shift:=xlUp
    wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
End Sub

Sub FootPrintNew()
    Dim iWhois, iNextWhois As Integer
    Dim sDomain As String
    
    Application.DisplayAlerts = False
    If WorksheetExists("DomainTMP") = True Then Sheets("DomainTMP").Delete
    Application.DisplayAlerts = True
    Sheets.Add.Name = "DomainTMP"
    Sheets("DomainTMP").Move After:=Sheets(5)
    
    For iWhois = 1 To 32767
        Sheets("DNSRecon").Select
        If Range("A" & iWhois).Value = "" Then Exit For
        If InStr(1, Range("A" & iWhois).Value, "dnsrecon -d ", vbTextCompare) > 0 Then
            sDomain = Right(Range("A" & iWhois).Value, Len(Range("A" & iWhois).Value) - InStr(1, Range("A" & iWhois).Value, "-d ", vbTextCompare) - 2)
            Sheets("DomainTMP").Select
            Call AddTempDomain(sDomain)
            Sheets("DNSRecon").Select
            iNextWhois = LocateNextDomain(iWhois + 1, "dnsrecon", sDomain)
        End If
        
        Call DNSReconLoop(iWhois, iNextWhois - 1, sDomain)
        Call MXReconLoop(iWhois, iNextWhois - 1, sDomain)
        
        iWhois = iNextWhois - 1
    Next iWhois
    
    Sheets("DomainTMP").Select
    For Each cell In Range("A:A")
        Sheets("DomainTMP").Select
        If cell.Value = "" Then Exit For
        Call DomainCleanup(cell.Value)
    Next cell
    
    Sheets("DomainTMP").Select
    For Each cell In Range("A:A")
        Sheets("DomainTMP").Select
        If cell.Value = "" Then Exit For
        Call WhoisNew(cell.Value)
    Next cell
    
    Application.DisplayAlerts = False
    If WorksheetExists("Footprinting") = True Then Sheets("Footprinting").Delete
    Application.DisplayAlerts = True
    Sheets.Add.Name = "Footprinting"
    Cells.Select
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 12
    End With
    Range("A1").Select
    Sheets("Footprinting").Move After:=Sheets(5)
    
    Sheets("DomainTMP").Select
    For Each cell In Range("A:A")
        Sheets("DomainTMP").Select
        If cell.Value = "" Then Exit For
        Call FootprintNewLoop(cell.Value, "A", "B")
        Call FootprintNewLoop(cell.Value, "C", "D")
        Call FootprintContactLoop(cell.Value)
    Next cell
    
    Sheets("DomainTMP").Select
    For Each cell In Range("A:A")
        Sheets("DomainTMP").Select
        If cell.Value = "" Then Exit For
        Call FootprintNewCleanupLoop(cell.Value)
    Next cell
    
    Sheets("DomainTMP").Select
    For Each cell In Range("A:A")
        If cell.Value = "" Then Exit For
        Application.DisplayAlerts = False
        If WorksheetExists(cell.Value) = True Then Sheets(cell.Value).Delete
        Application.DisplayAlerts = True
    Next cell
    
    Application.DisplayAlerts = False
    If WorksheetExists("DomainTMP") = True Then Sheets("DomainTMP").Delete
    Application.DisplayAlerts = True
    
    Sheets("Footprinting").Select
End Sub

Private Sub FootprintNewCleanupLoop(ByVal sDomain As String)
    Sheets("Footprinting").Select
    For Each cell In Range("A:A")
        If cell.Value = "" And Range("A" & cell.Row + 1).Value = "" Then Exit For
        If cell.Value = sDomain Then
            If cell.Row <> 1 Then
                Rows(cell.Row & ":" & cell.Row).Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Exit For
            End If
        End If
    Next cell
End Sub

Private Sub FootprintContactLoop(ByVal sDomain As String)
    Dim a, b, c, d, e, f As String
    Dim x As Integer
    Sheets(sDomain).Select
    a = Range("E1").Value
    b = Range("F1").Value
    c = Range("E2").Value
    d = Range("F2").Value
    e = Range("E3").Value
    f = Range("F3").Value
    
    Sheets("Footprinting").Select
    For Each cell In Range("A:A")
        If cell.Value = "" Then
            x = cell.Row
            Exit For
        End If
    Next cell
    
    If a <> "" Then
        Range("A" & x).Value = "Registrant Name"
        Range("B" & x).Value = a
        x = x + 1
    End If
    If d <> "" Then
        Range("A" & x).Value = "Registrant Email"
        Range("B" & x).Value = b
        x = x + 1
    End If
    If b <> "" Then
        Range("A" & x).Value = "Admin Name"
        Range("B" & x).Value = c
        x = x + 1
    End If
    If e <> "" Then
        Range("A" & x).Value = "Admin Email"
        Range("B" & x).Value = d
        x = x + 1
    End If
    If c <> "" Then
        Range("A" & x).Value = "Tech Name"
        Range("B" & x).Value = e
        x = x + 1
    End If
    If f <> "" Then
        Range("A" & x).Value = "Tech Email"
        Range("B" & x).Value = f
        x = x + 1
    End If
End Sub

Private Sub FootprintNewLoop(ByVal sDomain As String, sCol, sCol2)
    Sheets(sDomain).Select
    For Each cell In Range(sCol & ":" & sCol)
        If cell.Value = "" Then Exit For
        Sheets(sDomain).Select
        Call FootprintNewLoop2(sDomain, cell.Value, Range(sCol2 & cell.Row).Value, cell.Row, sCol)
    Next cell
End Sub

Private Sub FootprintNewLoop2(ByVal sDomain As String, ByVal x As String, ByVal z As String, ByVal iRow As Integer, ByVal sCol As String)
    Sheets("Footprinting").Select
    Dim y As Integer
    y = 1
    For Each cell In Range("A:A")
        If cell.Value = "" Then Exit For
        y = y + 1
    Next cell
    
    For y = y To 32767
        If iRow = 1 And sCol = "A" Then
            Range("A" & cell.Row).Value = sDomain
            'Range("B" & cell.Row).Value = sDomain
            y = y + 1
        End If
    
        If x <> "" Then
            Range("A" & y).Value = x
            Range("B" & y).Value = z
        End If
        Exit For
    Next y
End Sub

Private Sub WhoisNew(ByVal sDomain As String)
    Dim iWhois, iNextWhois As Integer
    Sheets("Whois").Select
    For iWhois = 1 To 32767
        If Range("A" & iWhois).Value = "" And Range("A" & iWhois + 1).Value = "" Then Exit For
        If InStr(1, Range("A" & iWhois).Value, "whois " & sDomain, vbTextCompare) > 0 Then
            iNextWhois = LocateNextDomain(iWhois + 1, "whois " & sDomain, sDomain)
            Call WhoisNewGather(sDomain, iWhois, iNextWhois)
            Exit For
        End If
    Next iWhois
End Sub

Private Sub WhoisNewGather(ByVal sDomain As String, ByVal iWhois As Integer, ByVal iNextWhois As Integer)
    Dim iAdminEmail, iTechEmail, iRegEmail, iAdminName, iTechName, iRegName
    Dim sAdminEmail, sTechEmail, sRegEmail, sAdminName, sTechName, sRegName
    Sheets("Whois").Select
    sAdminEmail = ""
    sTechEmail = ""
    sRegEmail = ""
    sAdminName = ""
    sTechName = ""
    sRegName = ""
    iAdminName = GetLocation(iWhois + 1, iNextWhois, "A", "Admin Name: ")
    iTechName = GetLocation(iWhois + 1, iNextWhois, "A", "Tech Name: ")
    iRegName = GetLocation(iWhois + 1, iNextWhois, "A", "Registrant Name: ")
    iAdminEmail = GetLocation(iWhois + 1, iNextWhois, "A", "Admin Email: ")
    iTechEmail = GetLocation(iWhois + 1, iNextWhois, "A", "Tech Email: ")
    iRegEmail = GetLocation(iWhois + 1, iNextWhois, "A", "Registrant Email: ")
    If iAdminName <> 0 Then
        sAdminName = Right(Range("A" & iAdminName).Value, Len(Range("A" & iAdminName).Value) - Len("Admin Name: "))
    End If
    If iTechName <> 0 Then
        sTechName = Right(Range("A" & iTechName).Value, Len(Range("A" & iTechName).Value) - Len("Tech Name: "))
    End If
    If iRegName <> 0 Then
        sRegName = Right(Range("A" & iRegName).Value, Len(Range("A" & iRegName).Value) - Len("Registrant Name: "))
    End If
    If iAdminEmail <> 0 Then
        sAdminEmail = Right(Range("A" & iAdminEmail).Value, Len(Range("A" & iAdminEmail).Value) - Len("Admin Email: "))
    End If
    If iTechEmail <> 0 Then
        sTechEmail = Right(Range("A" & iTechEmail).Value, Len(Range("A" & iTechEmail).Value) - Len("Tech Email: "))
    End If
    If iRegEmail <> 0 Then
        sRegEmail = Right(Range("A" & iRegEmail).Value, Len(Range("A" & iRegEmail).Value) - Len("Registrant Email: "))
    End If
     
    Sheets(sDomain).Select
    Range("E1").Value = sRegName
    Range("E2").Value = sAdminName
    Range("E3").Value = sTechName
    Range("F1").Value = sRegEmail
    Range("F2").Value = sAdminEmail
    Range("F3").Value = sTechEmail
End Sub

Private Sub DomainCleanup(ByVal sDomain As String)
    Sheets(sDomain).Select
    Dim x As Integer
    Dim y, z As String
    For x = 1 To 32767
        If Range("A" & x).Value = "" Then Exit For
        y = Range("A" & x).Value
        z = Range("B" & x).Value
        Call DomainCleanupLoop(x + 1, y, z)
    Next x
End Sub

Private Sub DomainCleanupLoop(ByVal x As Integer, ByVal y As String, ByVal z As String)
    For x = x To 32767
        If Range("A" & x).Value = "" Then Exit For
        If Range("A" & x).Value = y Then
            If Range("B" & x).Value = z Then
                Range("A" & x & ":B" & x).Select
                Selection.Delete Shift:=xlUp
                x = x - 1
            End If
        End If
    Next x
End Sub

Private Sub MXReconLoop(ByVal iWhois As Integer, ByVal iNextWhois As Integer, ByVal sDomain As String)
    For iWhois = iWhois To iNextWhois
        Sheets("DNSRecon").Select
        If InStr(1, Range("A" & iWhois).Value, "MX ", vbTextCompare) > 0 Then
            If InStr(1, Range("A" & iWhois).Value, "TXT", vbTextCompare) > 0 Then
            Else
                Dim x, y As String
                Dim z As Integer
                z = InStrRev(Range("A" & iWhois).Value, " ", , vbTextCompare)
                x = LCase(Trim(Mid(Range("A" & iWhois).Value, InStr(1, Range("A" & iWhois).Value, "MX ", vbTextCompare) + 3, z - InStr(1, Range("A" & iWhois).Value, "MX ", vbTextCompare) - 3)))
                y = Trim(Right(Range("A" & iWhois).Value, Len(Range("A" & iWhois).Value) - InStr(1, Range("A" & iWhois).Value, x, vbTextCompare) - Len(x)))
                Call ReconLoop2(sDomain, x, y, "C", "D")
            End If
        End If
    Next iWhois
End Sub

Private Sub DNSReconLoop(ByVal iWhois As Integer, ByVal iNextWhois As Integer, ByVal sDomain As String)
    For iWhois = iWhois To iNextWhois
        Sheets("DNSRecon").Select
        If InStr(1, Range("A" & iWhois).Value, "NS ", vbTextCompare) > 0 Then
            If InStr(1, Range("A" & iWhois).Value, "Trying", vbTextCompare) > 0 Then
            ElseIf InStr(1, Range("A" & iWhois).Value, "duplicate", vbTextCompare) > 0 Then
            ElseIf InStr(1, Range("A" & iWhois).Value, "Servers found", vbTextCompare) > 0 Then
            ElseIf InStr(1, Range("A" & iWhois).Value, "Resolving", vbTextCompare) > 0 Then
            Else
                Dim x, y As String
                Dim z As Integer
                z = InStrRev(Range("A" & iWhois).Value, " ", , vbTextCompare)
                x = LCase(Trim(Mid(Range("A" & iWhois).Value, InStr(1, Range("A" & iWhois).Value, "NS ", vbTextCompare) + 3, z - InStr(1, Range("A" & iWhois).Value, "NS ", vbTextCompare) - 3)))
                y = Trim(Right(Range("A" & iWhois).Value, Len(Range("A" & iWhois).Value) - InStr(1, Range("A" & iWhois).Value, x, vbTextCompare) - Len(x)))
                Call ReconLoop2(sDomain, x, y, "A", "B")
            End If
        End If
    Next iWhois
End Sub

Private Sub ReconLoop2(ByVal sDomain As String, ByVal x As String, ByVal y As String, ByVal sCol1 As String, ByVal sCol2 As String)
    Sheets(sDomain).Select
    For Each cell In Range(sCol1 & ":" & sCol1)
        If cell.Value = "" Then
            Range(sCol1 & cell.Row).Value = x
            Range(sCol2 & cell.Row).Value = y
            Exit For
        End If
    Next cell
End Sub

Private Sub AddTempDomain(ByVal sDomain As String)
    For Each cell In Range("A:A")
        If cell.Value = "" Then
            Range("A" & cell.Row).Value = sDomain
            Application.DisplayAlerts = False
            If WorksheetExists(sDomain) = True Then Sheets(sDomain).Delete
            Application.DisplayAlerts = True
            Sheets.Add.Name = sDomain
            Sheets(sDomain).Move After:=Sheets(5)
            Exit For
        End If
    Next cell
End Sub

Private Function LocateNextDomain(ByVal x As Integer, ByVal y As String, ByVal sDomain As String) As Integer
    If y = "dnsrecon" Then
        For x = x To 32767
            If Range("A" & x).Value = "" Then
                LocateNextDomain = Range("A" & x).Row
                Exit For
            End If
            If InStr(1, Range("A" & x).Value, "dnsrecon -d ", vbTextCompare) > 0 Then
                LocateNextDomain = Range("A" & x).Row
                Exit For
            End If
        Next x
    ElseIf y = "whois " & sDomain Then
        For x = x To 32767
            If Range("A" & x).Value = "" And Range("A" & x + 1).Value = "" And Range("A" & x + 2).Value = "" Then
                LocateNextDomain = Range("A" & x).Row
                Exit For
            End If
            If InStr(1, Range("A" & x).Value, "# whois ", vbTextCompare) > 0 Then
                LocateNextDomain = Range("A" & x).Row
                Exit For
            End If
        Next x
    End If
End Function

Sub Footprinting()
    Dim iWhois, iNextWhois, iAdminEmail, iTechEmail, iDNS, iNextDNS, iNS, iRoot, a, b, x As Integer
    Dim sAdminEmail, sTechEmail, sDNS, sDNSIP, sDomain, sCPU, sVal As String
    Dim boolDNS As Boolean
    iTechEmail = 0
    iAdminEmail = 0
    sAdminEmail = ""
    sTechEmail = ""
    iFootprint = 1
    Sheets("DNSRecon").Select
    Call GetDNSEOF
    
    'Get computer name
    If InStr(1, Range("A1"), "root@", vbTextCompare) Then
        iRoot = 1
    Else
        iRoot = GetLocation(1, 32767, "A", "root@")
    End If
    a = InStr(1, Range("A" & iRoot).Value, "@", vbTextCompare)
    b = InStr(1, Range("A" & iRoot).Value, ":", vbTextCompare)
    sCPU = Mid(Range("A" & iRoot).Value, a + 1, b - a - 1)
    
    For iWhois = 1 To 32767
        Sheets("Whois").Select
        sDomain = ""
        'Check for EOF
        If Range("A" & iWhois).Value = "" And Range("A" & iWhois + 1).Value = "" And Range("A" & iWhois + 2).Value = "" Then Exit For
        'Locate current whois record
        If Left(Range("A" & iWhois).Value, 15 + Len(sCPU)) = "root@" & sCPU & ":~# whois " Then
            sDomain = Right(Range("A" & iWhois).Value, Len(Range("A" & iWhois).Value) - 15 - Len(sCPU))
            'Locate next whois record
            iNextWhois = GetLocation(iWhois + 1, 32767, "A", "root@" & sCPU & ":~# whois ")
        End If
        
        If sDomain <> "" Then
            'Get registered emails if applicable
            sAdminEmail = ""
            sTechEmail = ""
            iAdminEmail = GetLocation(iWhois + 1, iNextWhois, "A", "Admin Email: ")
            iTechEmail = GetLocation(iWhois + 1, iNextWhois, "A", "Tech Email: ")
            If iAdminEmail <> 0 Then
                sAdminEmail = Right(Range("A" & iAdminEmail).Value, Len(Range("A" & iAdminEmail).Value) - 13)
            End If
            If iTechEmail <> 0 Then
                sTechEmail = Right(Range("A" & iTechEmail).Value, Len(Range("A" & iTechEmail).Value) - 13)
            End If
            
            'Write domain and registration to table
            Sheets("Footprinting").Select
            Range("A" & iFootprint).Value = sDomain
            Range("E" & iFootprint).Value = "FP01"
            iFootprint = iFootprint + 1
            Range("B" & iFootprint).Value = "Contact Information"
            iFootprint = iFootprint + 1
            Range("C" & iFootprint).Value = "Administrative Contact"
            If sAdminEmail <> "" Then
                Range("D" & iFootprint).Value = sAdminEmail
            Else
                Range("D" & iFootprint).Value = "N/A"
            End If
            Range("E" & iFootprint).Value = "FP02"
            iFootprint = iFootprint + 1
            Range("C" & iFootprint).Value = "Technical Contact"
            If sTechEmail <> "" Then
                Range("D" & iFootprint).Value = sTechEmail
            Else
                Range("D" & iFootprint).Value = "N/A"
            End If
            Range("E" & iFootprint).Value = "FP02"
            iFootprint = iFootprint + 1
                
            'Get DNS info
            Sheets("DNSRecon").Select
            iDNS = 0
            iNextDNS = 0
            x = 0
            iDNS = GetLocation(1, iDNSEOF, "A", "root@" & sCPU & ":~# dnsrecon -d " & sDomain)
            iNextDNS = GetLocation(iDNS + 1, iDNSEOF, "A", "root@" & sCPU & ":~# dnsrecon")
            If iNextDNS = 0 Then iNextDNS = iDNSEOF + 1
            boolDNS = False
            iAddDNS = 0
            For x = iDNS + 1 To iNextDNS - 1
                Sheets("DNSRecon").Select
                iAddDNS = 0
                iNS = GetLocation(x, iNextDNS, "A", " NS ")
                sVal = Range("A" & x).Value
                If iNS <> 0 Then
                    sDNS = Left(Mid(Range("A" & iNS).Value, 10, Len(Range("A" & iNS).Value) - 10), InStr(1, Mid(Range("A" & iNS).Value, 10, Len(Range("A" & iNS)) - 10), " ", vbTextCompare) - 1)
                    Sheets("Footprinting").Select
                    If sDNS <> "" And boolDNS = False Then
                        Range("B" & iFootprint).Value = "DNS Server Information"
                        iFootprint = iFootprint + 1
                        boolDNS = True
                    End If
                    If boolDNS = True Then
                        'If InStr(1, sVal, sDNS, vbTextCompare) > 0 And InStr(1, sVal, ":", vbTextCompare) = 0 Then
                            Call GetIPs(x, "A", sDNS, iNextDNS)
                        'End If
                        If Range("D" & iFootprint).Value <> "" Then
                            Range("C" & iFootprint).Value = sDNS
                            Range("E" & iFootprint).Value = "FP02"
                            iFootprint = iFootprint + 1
                        End If
                    End If
                    x = iNS + iAddDNS
                End If
            Next x
            
            'Get MX info
            Sheets("DNSRecon").Select
            x = 0
            boolDNS = False
            iAddDNS = 0
            For x = iDNS + 1 To iNextDNS - 1
                Sheets("DNSRecon").Select
                iAddDNS = 0
                iNS = GetLocation(x, iNextDNS, "A", " MX ")
                sVal = Range("A" & x).Value
                If iNS <> 0 Then
                    sDNS = Left(Mid(Range("A" & iNS).Value, 10, Len(Range("A" & iNS).Value) - 10), InStr(1, Mid(Range("A" & iNS).Value, 10, Len(Range("A" & iNS)) - 10), " ", vbTextCompare) - 1) 'update
                    Sheets("Footprinting").Select
                    If sDNS <> "" And boolDNS = False Then
                        Range("B" & iFootprint).Value = "Mail Server Information"
                        iFootprint = iFootprint + 1
                        boolDNS = True
                    End If
                    If boolDNS = True Then
                        'If InStr(1, sVal, sDNS, vbTextCompare) > 0 And InStr(1, sVal, ":", vbTextCompare) = 0 Then
                            Call GetIPs(x, "A", sDNS, iNextDNS)
                        'End If
                        If Range("D" & iFootprint).Value <> "" Then
                            Range("C" & iFootprint).Value = sDNS
                            Range("E" & iFootprint).Value = "FP02"
                            iFootprint = iFootprint + 1
                        End If
                    End If
                    x = iNS + iAddDNS
                End If
            Next x
        End If
    Next iWhois
    Sheets("Footprinting").Select
End Sub

Private Sub GetIPs(ByVal x As Integer, ByVal sCol As String, ByVal sDNS As String, ByVal iNextDNS As Integer)
    Dim sIPs As String
    Dim boolIPs
    boolIPs = False
    For x = x To iNextDNS - 1
        Sheets("DNSRecon").Select
        If Range(sCol & x).Value = "" And Range(sCol & x + 1).Value = "" And Range(sCol & x + 2).Value = "" Then
            x = iNextDNS
        End If
        If InStr(1, Range(sCol & x).Value, sDNS, vbTextCompare) > 0 And InStr(1, Range(sCol & x).Value, " SOA ", vbTextCompare) = 0 Then
            sIPs = Right(Range(sCol & x).Value, Len(Range(sCol & x).Value) - InStr(10, Range(sCol & x).Value, " ", vbTextCompare))
            If InStr(1, sIPs, ":", vbTextCompare) > 0 Then
            Else
                Sheets("Footprinting").Select
                If boolIPs = True Then sIPs = vbCrLf & sIPs
                If InStr(1, Range("D" & iFootprint).Value, Trim(sIPs), vbTextCompare) = 0 Then Range("D" & iFootprint).Value = Range("D" & iFootprint).Value & sIPs
                boolIPs = True
            End If
            iAddDNS = iAddDNS + 1
        End If
    Next x
    Sheets("Footprinting").Select
    iAddDNS = iAddDNS - 1
End Sub

Private Sub GetDNSEOF()
    For Each cell In Range("A:A")
        If cell.Value = "" Then Exit For
        iDNSEOF = iDNSEOF + 1
    Next cell
End Sub

Private Function GetLocation(ByVal iStartRow As Integer, ByVal iEndRow As Integer, ByVal sCol As String, ByVal sSearch As String)
    For iStartRow = iStartRow To iEndRow - 1
        If Range(sCol & iStartRow).Value = "" And Range(sCol & iStartRow + 1).Value = "" And Range(sCol & iStartRow + 2).Value = "" Then
            GetLocation = 0
            Exit For
        End If
        If InStr(1, Range(sCol & iStartRow).Value, sSearch, vbTextCompare) > 0 And InStr(1, Range(sCol & iStartRow).Value, ":", vbTextCompare) > 0 Then
            GetLocation = iStartRow
            Exit For
        End If
        GetLocation = 0
    Next iStartRow
End Function

Sub ImportNessusFile()
    Dim x, y As String
    y = Range("D1").Value
    
    If InStr(Len(y), y, "\", vbTextCompare) = 0 Then y = y & "\"
    y = y & Range("D3").Value
    ChDrive y
    ChDir y
    x = Application.GetOpenFilename("Nessus Files (*.nessus), *.nessus", , "Select Nessus File", "Import", False)
    
    If InStr(1, x, "/", vbTextCompare) <> 0 Then
        x = ConvertPath(x)
    Else
        x = x
    End If

    Application.DisplayAlerts = False
    If WorksheetExists("Nessus") = True Then Sheets("Nessus").Delete
    Application.DisplayAlerts = True
    Sheets.Add.Name = "Nessus"
    Sheets("Nessus").Move After:=Sheets(5)
    Sheets("Nessus").Select
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & x _
        , Destination:=Range("$A$1"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(2)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Sheets("Control").Select
End Sub

Sub ImportWhoisFile()
    Dim x, y As String
    y = Range("D1").Value
    
    If InStr(Len(y), y, "\", vbTextCompare) = 0 Then y = y & "\"
    y = y & Range("D3").Value
    ChDrive y
    ChDir y
    x = Application.GetOpenFilename("Text Files (*.txt), *.txt", , "Select Whois File", "Import", False)
    
    If InStr(1, x, "/", vbTextCompare) <> 0 Then
        x = ConvertPath(x)
    Else
        x = x
    End If

    Application.DisplayAlerts = False
    If WorksheetExists("Whois") = True Then Sheets("Whois").Delete
    Application.DisplayAlerts = True
    Sheets.Add.Name = "Whois"
    Sheets("Whois").Move After:=Sheets(5)
    Sheets("Whois").Select
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & x _
        , Destination:=Range("$A$1"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(2)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Sheets("Control").Select
End Sub

Sub ImportDNSReconFile()
    Dim x, y As String
    y = Range("D1").Value
    
    If InStr(Len(y), y, "\", vbTextCompare) = 0 Then y = y & "\"
    y = y & Range("D3").Value
    ChDrive y
    ChDir y
    x = Application.GetOpenFilename("Text Files (*.txt), *.txt", , "Select DNSRecon File", "Import", False)
    
    If InStr(1, x, "/", vbTextCompare) <> 0 Then
        x = ConvertPath(x)
    Else
        x = x
    End If

    Application.DisplayAlerts = False
    If WorksheetExists("DNSRecon") = True Then Sheets("DNSRecon").Delete
    Application.DisplayAlerts = True
    Sheets.Add.Name = "DNSRecon"
    Sheets("DNSRecon").Move After:=Sheets(6)
    Sheets("DNSRecon").Select
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & x _
        , Destination:=Range("$A$1"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(2)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Sheets("Control").Select
End Sub

Function FileSelectBox(ByRef FileType As String, Optional ByVal DefaultDir As String) As String
    Dim a As Object, FileName As String, varFile As Variant
    Set a = Application.FileDialog(msoFileDialogFilePicker)
    With a
        .AllowMultiSelect = False
        .Title = "Select Nessus File to Import"
        .Filters.Clear
        .Filters.Add "Nessus Files", FileType
        If Not IsMissing(DefaultDir) And DefaultDir <> "" Then .InitialFileName = DefaultDir
        If .Show = True Then
            For Each varFile In .SelectedItems
                FileSelectBox = varFile
            Next varFile
        End If
    End With
End Function

Sub WriteIPTReport()
    If Range("d2").Value = "" Then
        MsgBox "Please specify a working folder in step 1."
        Exit Sub
    End If
    Call WriteReport("IPT")
'Commented out calling the cover letters, they are in the new report template BB
'    Call WriteLetters("Draft Report Cover Letter.doc")
'    Call WriteLetters("Final Report Cover Letter.doc")
    Call WriteChecklists("MCS0500 Consulting Report Typing Checklist.xlsm", "Internal Penetration Testing", Range("A12").Value)
    Call WriteChecklists("MCS0501 Consulting Report Release Checklist.xlsm", "Internal Penetration Testing", Range("A6").Value)
    If Environ("username") = "ddennis" Then Call WriteChecklists("RM0620 Review Questionnaire - Cybersecurity.xlsm", "", "")
    MsgBox "Report complete.", , "Report complete"
End Sub

Sub WriteEPTReport()
    If Range("d2").Value = "" Then
        MsgBox "Please specify a working folder in step 1."
        Exit Sub
    End If
    Call WriteReport("EPT")
'Commented out calling the cover letters, they are in the new report template BB
'    Call WriteLetters("Draft Report Cover Letter.doc")
'    Call WriteLetters("Final Report Cover Letter.doc")
    Call WriteChecklists("MCS0500 Consulting Report Typing Checklist.xlsm", "External Penetration Testing", Range("A12").Value)
    Call WriteChecklists("MCS0501 Consulting Report Release Checklist.xlsm", "External Penetration Testing", Range("A6").Value)
    If Environ("username") = "ddennis" Then Call WriteChecklists("RM0620 Review Questionnaire - Cybersecurity.xlsm", "", "")
    MsgBox "Report complete.", , "Report complete"
End Sub

Sub WriteEVSReport()
    If Range("d2").Value = "" Then
        MsgBox "Please specify a working folder in step 1."
        Exit Sub
    End If
    Call WriteReport("EVS")
'Commented out calling the cover letters, they are in the new report template BB
'    Call WriteLetters("Draft Report Cover Letter.doc")
'    Call WriteLetters("Final Report Cover Letter.doc")
    Call WriteChecklists("MCS0500 Consulting Report Typing Checklist.xlsm", "External Vulnerability Scan", Range("A12").Value)
    Call WriteChecklists("MCS0501 Consulting Report Release Checklist.xlsm", "External Vulnerability Scan", Range("A6").Value)
    If Environ("username") = "ddennis" Then Call WriteChecklists("RM0620 Review Questionnaire - Cybersecurity.xlsm", "", "")
    MsgBox "Report complete.", , "Report complete"
End Sub

Sub WriteSEReport()
    If Range("d2").Value = "" Then
        MsgBox "Please specify a working folder in step 1."
        Exit Sub
    End If
    Call WriteReport("SE")
'Commented out calling the cover letters, they are in the new report template BB
'    Call WriteLetters("Draft Report Cover Letter.doc")
'    Call WriteLetters("Final Report Cover Letter.doc")
    Call WriteChecklists("MCS0500 Consulting Report Typing Checklist.xlsm", "Social Engineering", Range("A12").Value)
    Call WriteChecklists("MCS0501 Consulting Report Release Checklist.xlsm", "Social Engineering", Range("A6").Value)
    If Environ("username") = "ddennis" Then Call WriteChecklists("RM0620 Review Questionnaire - Cybersecurity.xlsm", "", "")
    MsgBox "Report complete.", , "Report complete"
End Sub

Sub WritePAReport()
    If Range("d2").Value = "" Then
        MsgBox "Please specify a working folder in step 1."
        Exit Sub
    End If
    Call WriteReport("PA")
'Commented out calling the cover letters, they are in the new report template BB
'    Call WriteLetters("Draft Report Cover Letter.doc")
'    Call WriteLetters("Final Report Cover Letter.doc")
    Call WriteChecklists("MCS0500 Consulting Report Typing Checklist.xlsm", "Other:", Range("A12").Value)
    Call WriteChecklists("MCS0501 Consulting Report Release Checklist.xlsm", "Other:", Range("A6").Value)
    If Environ("username") = "ddennis" Then Call WriteChecklists("RM0620 Review Questionnaire - Cybersecurity.xlsm", "", "")
    MsgBox "Report complete.", , "Report complete"
End Sub

Private Function GetUsername()
    Dim objAD, objUser As Object
    Set objAD = CreateObject("ADSystemInfo")
    Set objUser = GetObject("LDAP://" & objAD.UserName)
    sFullUsername = objUser.DisplayName
    GetUsername = sFullUsername
End Function

Private Sub WriteChecklists(ByVal sEng As String, ByVal sCyber As String, ByVal sDate As String)
    Sheets("Control").Select
    Dim FilePath As String
    Dim xlWorkbook As Workbook
    FilePath = sBase1 & sBase3 & sSlash1 & sEng
    Call CopyFiles(sEng, FilePath)
    Set xlWorkbook = Workbooks.Open(FilePath)
    
    If sEng = "MCS0500 Consulting Report Typing Checklist.xlsm" Then
        xlWorkbook.Worksheets("Printing and Delivery").Range("Deliverable").Value = "Cybersecurity"
        xlWorkbook.Worksheets("Printing and Delivery").Range("Cyber").Value = sCyber
        xlWorkbook.Worksheets("Printing and Delivery").Range("C13").Value = "RM40"
        xlWorkbook.Worksheets("Printing and Delivery").Range("draft").Value = "Yes"
        If ThisWorkbook.Sheets("Control").chkFinalOnly.Value = False Then
            If sDate <> "" Then xlWorkbook.Worksheets("Printing and Delivery").Range("F19").Value = Format(CDate(sDate), "mm/dd/yyyy")
        Else
            If sDate <> "" Then xlWorkbook.Worksheets("Printing and Delivery").Range("F20").Value = Format(CDate(sDate), "mm/dd/yyyy")
        End If
        xlWorkbook.Worksheets("Printing and Delivery").Range("D25").Value = GetUsername
        xlWorkbook.Worksheets("Printing and Delivery").Range("B33").Value = 1
        xlWorkbook.Worksheets("Printing and Delivery").CheckBoxes("Check Box 78").Value = True
        xlWorkbook.Worksheets("Printing and Delivery").Range("L54").Value = "Send PDF to " & Environ("username") & "@bkd.com"
    ElseIf sEng = "MCS0501 Consulting Report Release Checklist.xlsm" Then
        xlWorkbook.Worksheets("Report Release").Range("Deliverable").Value = "Cybersecurity"
        xlWorkbook.Worksheets("Report Release").Range("Cyber").Value = sCyber
        xlWorkbook.Worksheets("Report Release").Range("E12").Value = "RM40"
        If ThisWorkbook.Sheets("Control").chkFinalOnly.Value = True Then If sDate <> "" Then xlWorkbook.Worksheets("Report Release").Range("G16").Value = Format(CDate(sDate), "mm/dd/yyyy")
        xlWorkbook.Worksheets("Report Release").Range("E18").Value = GetUsername
        If sDate <> "" Then xlWorkbook.Worksheets("Report Release").Range("J18").Value = Format(CDate(sDate), "mm/dd/yyyy")
        xlWorkbook.Worksheets("Report Release").Range("Eval").Value = "No"
    End If
    xlWorkbook.Close (True)
End Sub

Private Sub WriteLetters(ByVal sEng As String)
    Sheets("Control").Select
    Dim FilePath As String
    'Dim x As Integer
    'Dim currentRange As Range
    'filepath = Range("d2").Value & "\" & sEng
    FilePath = sBase1 & sBase3 & sSlash1 & sEng
    Call CopyFiles(sEng, FilePath)
    
    Dim wdApp As Object
    Dim wdDoc As Object
    
    Set wdApp = CreateObject("word.application")
    wdApp.Visible = True
    Set wdoc = wdApp.Documents.Open(FilePath)
    Set WordContent = wdoc.Content
    wdoc.TrackRevisions = True
    
    For Each cell In Range("B2:B32767")
        If cell.Value = "" Then Exit For
        If Range("B" & cell.Row).Value = "{today}" Then
            With WordContent.Find
                .Text = "{today}"
                '.Replacement.Text = Format(Range("A" & cell.Row).Value, "MMMM DD, YYYY")
                .Replacement.Text = Range("A" & cell.Row).Text
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
        End If
        If Range("B" & cell.Row).Value = "{title}" Then
            If Range("A" & cell.Row).Value = "" Then
                With WordContent.Find
                    .Text = "{title}"
                    .Replacement.Text = ""
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                End With
            Else
                With WordContent.Find
                    .Text = "{title}"
                    .Replacement.Text = ", " & Range("A" & cell.Row).Value
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                End With
            End If
        Else
            With WordContent.Find
                .Text = Range("B" & cell.Row).Value
                .Replacement.Text = Range("A" & cell.Row).Value
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
        End If
    Next cell
    wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
    
End Sub

Function GetFolder() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

Sub PickClientFolder()
    Dim x As String
    Dim y As Integer
    x = GetFolder
    
    If InStr(1, x, "\", vbTextCompare) Then
        y = InStrRev(x, "\", , vbTextCompare)
        x = Trim(Right(x, Len(x) - y))
    Else
        y = InStrRev(x, "/", , vbTextCompare)
        x = Trim(Right(x, Len(x) - y))
    End If
       
    Range("d3").Value = x
End Sub

Sub PickScreenshot1()
    Dim x, y As String
    y = Range("D1").Value
    
    If InStr(Len(y), y, "\", vbTextCompare) = 0 Then y = y & "\"
    y = y & Range("D3").Value
    ChDrive y
    ChDir y
    x = Application.GetOpenFilename("PNG Files (*.png), *.png", , "Select Image", "Import", False)
    
    If InStr(1, x, "/", vbTextCompare) <> 0 Then
        Range("H7").Value = ConvertPath(x)
    Else
        Range("H7").Value = x
    End If
    ActiveWindow.ScrollRow = 1
End Sub

Sub PickScreenshot2()
    Dim x, y As String
    y = Range("D1").Value
    
    If InStr(Len(y), y, "\", vbTextCompare) = 0 Then y = y & "\"
    y = y & Range("D3").Value
    ChDrive y
    ChDir y
    x = Application.GetOpenFilename("PNG Files (*.png), *.png", , "Select Image", "Import", False)
    
    If InStr(1, x, "/", vbTextCompare) <> 0 Then
        Range("H8").Value = ConvertPath(x)
    Else
        Range("H8").Value = x
    End If
    ActiveWindow.ScrollRow = 1
End Sub

Sub PickPhishingFile()
    Dim x, y, c, d As String
    Dim a, b, z As Integer
    y = Range("D1").Value
    
    If InStr(Len(y), y, "\", vbTextCompare) = 0 Then y = y & "\"
    y = y & Range("D3").Value
    ChDrive y
    ChDir y
    x = Application.GetOpenFilename("Phishing File (*.xlsm), *.xlsm", , "Select Phishing", "Import", False)
    
    If InStr(1, x, "/", vbTextCompare) <> 0 Then
        x = ConvertPath(x)
        Range("H5").Value = x
    Else
        Range("H5").Value = x
    End If
    
    If InStr(1, x, "\", vbTextCompare) Then
        z = InStrRev(x, "\", , vbTextCompare)
        y = Trim(Right(x, Len(x) - z))
    Else
        z = InStrRev(x, "/", , vbTextCompare)
        y = Trim(Right(x, Len(x) - z))
    End If
    
    Range("H14").Formula = "='" & Left(x, Len(x) - Len(y)) & "[" & y & "]Control'!$B$30"
    d = Range("H14").Value
    Range("H14").Value = d
    
    Range("H15").Formula = "='" & Left(x, Len(x) - Len(y)) & "[" & y & "]Control'!$B$49"
    d = Range("H15").Value
    Range("H15").Value = d
    
    Range("H20").Formula = "='" & Left(x, Len(x) - Len(y)) & "[" & y & "]Employee Info'!$E$1"
    d = Range("H20").Value
    Range("H20").Value = d
    
    Range("I20").Value = "='" & Left(x, Len(x) - Len(y)) & "[" & y & "]Employee Info'!$F$1"
    d = Range("I20").Value
    z = CInt(d)
    Range("I20").Value = ""

    Application.DisplayAlerts = False
    If WorksheetExists("Emails") = True Then Sheets("Emails").Delete
    Application.DisplayAlerts = True
    Sheets.Add.Name = "Emails"
    Sheets("Emails").Move Before:=Sheets(3)
    Sheets("Control").Select
    b = 1
    Sheets("Emails").Select
    For a = 2 To z + 1
        c = "='" & Left(x, Len(x) - Len(y)) & "[" & y & "]Employee Info'!$B$" & a
        Range("A" & b).Formula = c
        c = Range("A" & b).Value
        Range("A" & b).Value = c
        
        If Range("A" & b).Value = "" Then
            Exit For
        ElseIf InStr(1, Range("A" & b).Value, "@bkd.com", vbTextCompare) = 0 Then
            b = b + 1
        End If
    Next a
    Sheets("Control").Select
    ActiveWindow.ScrollRow = 1
End Sub

Function WorksheetExists(sName As String) As Boolean
    WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
End Function

Function ConvertPath(ByVal x As String) As String
    Dim y, z, a As String
    Dim i As Long
    y = Environ$("username")
    a = Range("D1").Value
    i = InStr(1, x, y, vbTextCompare)
    z = Right(x, Len(x) - i - Len(y) - 10)
    z = Replace(z, "/", "\", 1, , vbTextCompare)
    i = InStr(1, a, "\", vbTextCompare)
    i = Len(a) - i
    z = a & Right(z, Len(z) - i)
    ConvertPath = z
End Function

Sub PickEmailFile()
    Dim x, y As String
    Dim z As Integer
    y = Range("D1").Value
    
    If InStr(Len(y), y, "\", vbTextCompare) = 0 Then y = y & "\"
    y = y & Range("D3").Value
    ChDrive y
    ChDir y
    
    x = Application.GetOpenFilename("Phishing Email File (*.msg), *.msg", , "Select Phishing Email", "Import", False)

    If InStr(1, x, "/", vbTextCompare) <> 0 Then
        x = ConvertPath(x)
        Range("H6").Value = x
    Else
        Range("H6").Value = x
    End If

    Dim olApp As Object
    Set olApp = CreateObject("Outlook.Application")
    Dim mailDoc As Object
    Set mailDoc = olApp.CreateItem(olMailItem)
    
    Dim nam As Variant
    For Each nam In Array(x)
        Set mailDoc = olApp.session.OpenSharedItem(x)
        Range("H9").Value = Format(mailDoc.SentOn, "Long Date")
        Range("H10").Value = mailDoc.SenderName
        Range("H11").Value = mailDoc.SenderEmailAddress
        Range("H12").Value = mailDoc.Subject
        Range("H13").Value = mailDoc.Body
        Range("H13").Select
        Selection.WrapText = False
    Next nam
    ActiveWindow.ScrollRow = 1
End Sub

Sub ImportClicks()
    Dim x, y As String
    y = Range("D1").Value
    
    If InStr(Len(y), y, "\", vbTextCompare) = 0 Then y = y & "\"
    y = y & Range("D3").Value
    ChDrive y
    ChDir y
    x = Application.GetOpenFilename("TXT Files (*.txt), *.txt", , "Select Counter File", "Import", False)
    
    If InStr(1, x, "/", vbTextCompare) <> 0 Then
        x = ConvertPath(x)
    End If
    
    Dim TextFile As Integer
    Dim FilePath As String
    Dim FileContent As String
    
    'File Path of Text File
      FilePath = x
    
    'Determine the next file number available for use by the FileOpen function
      TextFile = FreeFile
    
    'Open the text file
      Open FilePath For Input As TextFile
    
    'Store file content inside a variable
      FileContent = Input(LOF(TextFile), TextFile)
    
    'Report Out Text File Contents
      Range("H16").Value = FileContent
    
    'Close Text File
      Close TextFile
End Sub

Sub ImportCredentials()
    Dim x, y As String
    Dim z As Integer
    y = Range("D1").Value
    
    If InStr(Len(y), y, "\", vbTextCompare) = 0 Then y = y & "\"
    y = y & Range("D3").Value
    ChDrive y
    ChDir y
    x = Application.GetOpenFilename("Credentials File (*.txt), *.txt", , "Select Credentials", "Import", False)
    If InStr(1, x, "/", vbTextCompare) <> 0 Then
        x = ConvertPath(x)
    End If
    
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("Credentials").Delete
    Application.DisplayAlerts = True
    
    Sheets.Add.Name = "Credentials"
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & x _
        , Destination:=Range("$A$1"))
        .Name = "credentials"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileOtherDelimiter = ""
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    z = 0
    For Each cell In Range("A:A")
        If cell.Value = "" Then Exit For
        z = z + 1
    Next cell
    
    Sheets("Control").Select
    Range("H17").Value = z
    Sheets("Credentials").Select
    Sheets("Credentials").Move Before:=Sheets(3)
    MsgBox "Please remove all invalid submissions before returning to the Control tab."
End Sub

Sub SetFindings()
    Sheets("Findings").Select
End Sub

Sub SetStrengthWeakness()
    Sheets("StrengthWeakness").Select
End Sub

Sub CustomMailMessage(ByVal sEng As String)
    Dim OutApp As Object
    Dim objOutlookMsg As Object
    Dim objOutlookRecip As Object
    Dim Recipients As Object
    
    Set OutApp = CreateObject("Outlook.Application")
    Set objOutlookMsg = OutApp.CreateItem(olMailItem)
    
    Set Recipients = objOutlookMsg.Recipients
    Set objOutlookRecip = Recipients.Add("claughridge@bkd.com")
    objOutlookRecip.Type = 1
    
    objOutlookMsg.Subject = "Binder Review - " & Range("A2").Value & " (" & Range("A11").Value & ") " & sEng & " Draft"
    objOutlookMsg.Display
    'objOutlookMsg.HTMLBody = "<html><body style=font-size:10pt;font-family:Arial>Chuck:" & "<br/><br/>" & _
        "This binder is ready to review:" & "<br/><br/>" & _
        "Pfx CFR: <strong>ITRS</strong>" & "<br/>" & _
        "Pfx Client: <strong>" & Range("A2").Value & " (" & Range("A11").Value & ") " & "</strong><br/>" & _
        "Pfx Binder ID: <strong>" & sEng & " " & Year(Range("A6").Value) & "</strong><br/>" & _
        "Pfx Binder tab: <strong>" & " 210 " & sEng & "</strong><br/>" & _
        "Report: <strong>Draft Report</strong>" & "<br/>" & _
        "Date Requested: <strong><span style='background:yellow;mso-highlight:yellow'>" & Format(Range("A12").Value, yyyy) & "</span></strong><br/>" & _
        objOutlookMsg.HTMLBody & "</body></html>"
    objOutlookMsg.HTMLBody = "<html><body style=font-size:10pt;font-family:Arial>Chuck:" & "<br/><br/>" & _
        "This binder is ready to review:" & "<br/><br/>" & _
        "Pfx CFR: <strong>ITRS</strong>" & "<br/>" & _
        "Pfx Client: <strong>" & Range("A2").Value & " (" & Range("A11").Value & ") " & "</strong><br/>" & _
        "Pfx Binder ID: <strong>" & " PT " & Year(Range("A6").Value) & "</strong><br/>" & _
        "Pfx Binder tab: <strong>" & " 210 " & sEng & "</strong><br/>" & _
        "Report: <strong>Draft Report</strong>" & "<br/>" & _
        "Date Requested: <strong><span style='background:yellow;mso-highlight:yellow'>" & Format(Range("A12").Value, yyyy) & "</span></strong><br/>" & _
        objOutlookMsg.HTMLBody & "</body></html>"
    objOutlookMsg.Importance = 1

    'Resolve each Recipient's name.
    For Each objOutlookRecip In objOutlookMsg.Recipients
      objOutlookRecip.Resolve
    Next
    'objOutlookMsg.Send
    objOutlookMsg.Display
    
    Set OutApp = Nothing
End Sub

Private Sub SetPageNumbers(ByRef wdoc As Object)
    Call FindReplace(wdoc, "i")
    Call FindReplace(wdoc, "ii")
    Call FindReplace(wdoc, "iii")
    Call FindReplace(wdoc, "iv")
    Call FindReplace(wdoc, "v")
    Call FindReplace(wdoc, "vi")
    Call FindReplace(wdoc, "vii")
    Call FindReplace(wdoc, "viii")
    Call FindReplace(wdoc, "ix")
    Call FindReplace(wdoc, "x")
End Sub

Private Sub FindReplace(ByRef wdoc As Object, ByVal x As String)
    Dim y As String
    wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
    wdoc.ActiveWindow.Selection.Find.Text = "{section " & x & "}"
    If wdoc.ActiveWindow.Selection.Find.Execute Then
        wdoc.ActiveWindow.Selection.TypeBackspace
        y = wdoc.ActiveWindow.Selection.Information(wdActiveEndAdjustedPageNumber)
        wdoc.ActiveWindow.Selection.HomeKey Unit:=wdStory
        wdoc.ActiveWindow.Selection.Find.Text = "{" & x & "}"
        wdoc.ActiveWindow.Selection.Find.Execute
        wdoc.ActiveWindow.Selection.TypeBackspace
        wdoc.ActiveWindow.Selection.TypeText Text:=y
    End If
End Sub

Sub SelectSummary()
    ActiveWindow.SmallScroll Down:=21
End Sub

