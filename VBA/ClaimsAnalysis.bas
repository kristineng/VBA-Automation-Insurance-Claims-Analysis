Attribute VB_Name = "modClaims"
Option Explicit
Sub CleanClaimsData()
    '----------------------------------------------------
    ' DECLARE VARIABLES
    '----------------------------------------------------
    Dim wsRaw As Worksheet
    Dim wsClean As Worksheet
    Dim wsQuality As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim missingAmount As Long
    Dim missingPolicyClaims As Long
    Dim duplicateClaims As Long
    Dim invalidAmount As Long
    Dim invalidDate As Long
    Dim qualityFlag As String
    Dim amountValue As String
    '----------------------------------------------------
    ' CONNECT VBA TO WORKSHEETS
    '----------------------------------------------------
    Set wsRaw = ThisWorkbook.Worksheets("Raw_Data")
    Set wsClean = ThisWorkbook.Worksheets("Clean_Data")
    Set wsQuality = ThisWorkbook.Worksheets("Data_Quality")

    lastRow = wsRaw.Cells(wsRaw.Rows.Count, "A").End(xlUp).Row
    wsClean.Cells.Clear
    wsRaw.Range("A1:J" & lastRow).Copy _
        Destination:=wsClean.Range("A1")
        
    wsClean.Cells(1, 11).Value = "claim_severity"
    wsClean.Cells(1, 12).Value = "data_quality_flag"
    '----------------------------------------------------
    ' CHECK EVERY CLAIM
    '----------------------------------------------------
    For i = 2 To lastRow
        qualityFlag = "OK"
        If Trim(CStr(wsClean.Cells(i, 8).Value)) = "" Then
            missingAmount = missingAmount + 1
            qualityFlag = "Missing Claim Amount"
        Else
            amountValue = CStr(wsClean.Cells(i, 8).Value)
            amountValue = Replace(amountValue, "$", "")
            'Remove commas
            amountValue = Replace(amountValue, ",", "")
            'Remove unnecessary spaces
            amountValue = Trim(amountValue)
            'Check whether it is a number
            If IsNumeric(amountValue) Then
                wsClean.Cells(i, 8).Value = CDbl(amountValue)
            Else
                invalidAmount = invalidAmount + 1
                If qualityFlag = "OK" Then
                    qualityFlag = "Invalid Claim Amount"
                End If
            End If
        End If
        '------------------------------------------------
        ' CHECK CLAIM DATE
        '------------------------------------------------
        If Trim(CStr(wsClean.Cells(i, 4).Value)) = "" Then
            invalidDate = invalidDate + 1
            If qualityFlag = "OK" Then
                qualityFlag = "Missing Claim Date"
            End If
        ElseIf IsDate(wsClean.Cells(i, 4).Value) Then
            wsClean.Cells(i, 4).Value = _
                CDate(wsClean.Cells(i, 4).Value)
        Else
            invalidDate = invalidDate + 1
            If qualityFlag = "OK" Then
                qualityFlag = "Invalid Date"
            End If
        End If
        '------------------------------------------------
        ' CHECK TOTAL POLICY CLAIMS
        '------------------------------------------------
        If Trim(CStr(wsClean.Cells(i, 9).Value)) = "" Then
            missingPolicyClaims = missingPolicyClaims + 1
            If qualityFlag = "OK" Then
                qualityFlag = "Missing Policy Claims"
            End If
        End If
        '------------------------------------------------
        ' CHECK DUPLICATE CLAIM IDs
        '------------------------------------------------
        If Application.WorksheetFunction.CountIf( _
            wsClean.Range("A$2:A$" & lastRow), _
            wsClean.Cells(i, 1).Value) > 1 Then
            duplicateClaims = duplicateClaims + 1
            If qualityFlag = "OK" Then
                qualityFlag = "Duplicate Claim ID"
            End If
        End If
        '------------------------------------------------
        ' CREATE CLAIM SEVERITY
        '------------------------------------------------
        If IsNumeric(wsClean.Cells(i, 8).Value) Then
            If wsClean.Cells(i, 8).Value < 5000 Then
                wsClean.Cells(i, 11).Value = "Low"
            ElseIf wsClean.Cells(i, 8).Value < 20000 Then
                wsClean.Cells(i, 11).Value = "Medium"
            ElseIf wsClean.Cells(i, 8).Value < 50000 Then
                wsClean.Cells(i, 11).Value = "High"
            Else
                wsClean.Cells(i, 11).Value = "Very High"
            End If
        Else
            wsClean.Cells(i, 11).Value = "Unknown"
        End If
        '------------------------------------------------
        ' INCLUDE DATA QUALITY FLAG
        '------------------------------------------------
        wsClean.Cells(i, 12).Value = qualityFlag
    Next i
    '------------------------------------------------
    ' FORMAT THE CLEANED DATA
    '------------------------------------------------
    wsClean.Range("D2:D" & lastRow).NumberFormat = "dd/mm/yyyy"
    wsClean.Range("H2:H" & lastRow).NumberFormat = "$#,##0.00"
    wsClean.Range("A1:L1").Font.Bold = True
    wsClean.Columns("A:L").AutoFit
    '------------------------------------------------
    ' 9. CREATE DATA QUALITY SUMMARY
    '------------------------------------------------
    wsQuality.Cells.Clear
    wsQuality.Range("A1").Value = "DATA QUALITY CHECK"
    wsQuality.Range("A1").Font.Bold = True
    wsQuality.Range("A3").Value = "Total Records"
    wsQuality.Range("B3").Value = lastRow - 1
    wsQuality.Range("A4").Value = "Missing Claim Amounts"
    wsQuality.Range("B4").Value = missingAmount
    wsQuality.Range("A5").Value = "Missing Policy Claim Counts"
    wsQuality.Range("B5").Value = missingPolicyClaims
    wsQuality.Range("A6").Value = "Duplicate Claim ID Records"
    wsQuality.Range("B6").Value = duplicateClaims
    wsQuality.Range("A7").Value = "Invalid Claim Amounts"
    wsQuality.Range("B7").Value = invalidAmount
    wsQuality.Range("A8").Value = "Invalid Dates"
    wsQuality.Range("B8").Value = invalidDate
    wsQuality.Columns("A:B").AutoFit

    MsgBox "Claims data cleaning completed successfully.", _
           vbInformation
End Sub

Sub ClaimsAnalysis()
    Dim wsClean As Worksheet
    Dim wsAnalysis As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim outputRow As Long
    Dim startRow As Long
    Dim totalClaims As Long
    Dim totalAmount As Double
    Dim averageAmount As Double
    Dim medianAmount As Double
    Dim maximumAmount As Double
    Dim minimumAmount As Double
    Dim fraudulentClaims As Long
    Dim nonFraudulentClaims As Long
    Dim fraudRate As Double
    Dim fraudulentAmount As Double
    Dim nonFraudulentAmount As Double
    Dim averageFraudulentClaim As Double
    Dim claimType As String
    Dim incidentCause As String
    Dim claimArea As String
    Dim policeReport As String
    Dim severity As String
    Dim customerID As String
    Dim claimYear As String
    Dim claimAmount As Double
    Dim dictType As Object
    Dim dictCause As Object
    Dim dictArea As Object
    Dim dictPolice As Object
    Dim dictSeverity As Object
    Dim dictCustomer As Object
    Dim dictYear As Object
    Dim key As Variant
    Dim multipleClaimCustomers As Long
    Set wsClean = ThisWorkbook.Worksheets("Clean_Data")
    Set wsAnalysis = ThisWorkbook.Worksheets("Analysis")
    lastRow = wsClean.Cells(wsClean.Rows.Count, "A").End(xlUp).Row

    Set dictType = CreateObject("Scripting.Dictionary")
    Set dictCause = CreateObject("Scripting.Dictionary")
    Set dictArea = CreateObject("Scripting.Dictionary")
    Set dictPolice = CreateObject("Scripting.Dictionary")
    Set dictSeverity = CreateObject("Scripting.Dictionary")
    Set dictCustomer = CreateObject("Scripting.Dictionary")
    Set dictYear = CreateObject("Scripting.Dictionary")
    wsAnalysis.Cells.Clear
    
    '----------------------------------------------------
    ' LOOP
    '----------------------------------------------------
    For i = 2 To lastRow
        '------------------------------------------------
        ' CLAIM AMOUNT / PORTFOLIO
        '------------------------------------------------
        If IsNumeric(wsClean.Cells(i, 8).Value) Then
            claimAmount = CDbl(wsClean.Cells(i, 8).Value)
            totalClaims = totalClaims + 1
            totalAmount = totalAmount + claimAmount
        End If
        '------------------------------------------------
        ' CLAIM TYPE
        '------------------------------------------------
        claimType = Trim(CStr(wsClean.Cells(i, 7).Value))
        If claimType = "" Then
            claimType = "Unknown"
        End If
        If Not dictType.Exists(claimType) Then
            dictType.Add claimType, Array(1, 0#)
        Else
            dictType(claimType) = Array( _
                dictType(claimType)(0) + 1, _
                dictType(claimType)(1))
        End If
        If IsNumeric(wsClean.Cells(i, 8).Value) Then
            dictType(claimType) = Array( _
                dictType(claimType)(0), _
                dictType(claimType)(1) + claimAmount)
        End If
        '------------------------------------------------
        ' INCIDENT CAUSE
        '------------------------------------------------
        incidentCause = Trim(CStr(wsClean.Cells(i, 3).Value))
        If incidentCause = "" Then
            incidentCause = "Unknown"
        End If
        If Not dictCause.Exists(incidentCause) Then
            dictCause.Add incidentCause, Array(1, 0#)
        Else
            dictCause(incidentCause) = Array( _
                dictCause(incidentCause)(0) + 1, _
                dictCause(incidentCause)(1))
        End If
        If IsNumeric(wsClean.Cells(i, 8).Value) Then
            dictCause(incidentCause) = Array( _
                dictCause(incidentCause)(0), _
                dictCause(incidentCause)(1) + claimAmount)
        End If
        '------------------------------------------------
        ' CLAIM AREA
        '------------------------------------------------
        claimArea = Trim(CStr(wsClean.Cells(i, 5).Value))
        If claimArea = "" Then
            claimArea = "Unknown"
        End If
        If Not dictArea.Exists(claimArea) Then
            dictArea.Add claimArea, Array(1, 0#)
        Else
            dictArea(claimArea) = Array( _
                dictArea(claimArea)(0) + 1, _
                dictArea(claimArea)(1))
        End If
        If IsNumeric(wsClean.Cells(i, 8).Value) Then
            dictArea(claimArea) = Array( _
                dictArea(claimArea)(0), _
                dictArea(claimArea)(1) + claimAmount)
        End If
        '------------------------------------------------
        ' POLICE REPORT
        '------------------------------------------------
        policeReport = Trim(CStr(wsClean.Cells(i, 6).Value))
        If policeReport = "" Then
            policeReport = "Unknown"
        End If
        If Not dictPolice.Exists(policeReport) Then
            dictPolice.Add policeReport, Array(1, 0#)
        Else
            dictPolice(policeReport) = Array( _
                dictPolice(policeReport)(0) + 1, _
                dictPolice(policeReport)(1))
        End If
        If IsNumeric(wsClean.Cells(i, 8).Value) Then
            dictPolice(policeReport) = Array( _
                dictPolice(policeReport)(0), _
                dictPolice(policeReport)(1) + claimAmount)
        End If
        '------------------------------------------------
        ' CLAIM SEVERITY
        '------------------------------------------------
        severity = Trim(CStr(wsClean.Cells(i, 11).Value))
        If severity = "" Then
            severity = "Unknown"
        End If
        If Not dictSeverity.Exists(severity) Then
            dictSeverity.Add severity, Array(1, 0#)
        Else
            dictSeverity(severity) = Array( _
                dictSeverity(severity)(0) + 1, _
                dictSeverity(severity)(1))
        End If
        If IsNumeric(wsClean.Cells(i, 8).Value) Then
            dictSeverity(severity) = Array( _
                dictSeverity(severity)(0), _
                dictSeverity(severity)(1) + claimAmount)
        End If
        '------------------------------------------------
        ' CUSTOMER
        '------------------------------------------------
        customerID = Trim(CStr(wsClean.Cells(i, 2).Value))
        If customerID <> "" Then
            If Not dictCustomer.Exists(customerID) Then
                dictCustomer.Add customerID, 1
            Else
                dictCustomer(customerID) = _
                    dictCustomer(customerID) + 1
            End If
        End If
        '------------------------------------------------
        ' TIME / YEAR
        '------------------------------------------------
        If IsDate(wsClean.Cells(i, 4).Value) Then
            claimYear = CStr(Year(wsClean.Cells(i, 4).Value))
            If Not dictYear.Exists(claimYear) Then
                dictYear.Add claimYear, Array(1, 0#)
            Else
                dictYear(claimYear) = Array( _
                    dictYear(claimYear)(0) + 1, _
                    dictYear(claimYear)(1))
            End If
            If IsNumeric(wsClean.Cells(i, 8).Value) Then
                dictYear(claimYear) = Array( _
                    dictYear(claimYear)(0), _
                    dictYear(claimYear)(1) + claimAmount)
            End If
        End If
        '------------------------------------------------
        ' FRAUD
        '------------------------------------------------
        If UCase(Trim(CStr(wsClean.Cells(i, 10).Value))) = "YES" Then
            fraudulentClaims = fraudulentClaims + 1
            If IsNumeric(wsClean.Cells(i, 8).Value) Then
                fraudulentAmount = _
                    fraudulentAmount + claimAmount
            End If
        Else
            nonFraudulentClaims = _
                nonFraudulentClaims + 1
            If IsNumeric(wsClean.Cells(i, 8).Value) Then
                nonFraudulentAmount = _
                    nonFraudulentAmount + claimAmount
            End If
        End If
    Next i
    '----------------------------------------------------
    ' CALCULATE PORTFOLIO METRICS
    '----------------------------------------------------
    If totalClaims > 0 Then
        averageAmount = totalAmount / totalClaims
        medianAmount = Application.WorksheetFunction.Median( _
            wsClean.Range("H2:H" & lastRow))
        maximumAmount = Application.WorksheetFunction.Max( _
            wsClean.Range("H2:H" & lastRow))
        minimumAmount = Application.WorksheetFunction.Min( _
            wsClean.Range("H2:H" & lastRow))
    End If
    If totalClaims > 0 Then
        fraudRate = fraudulentClaims / totalClaims
    End If
    If fraudulentClaims > 0 Then
        averageFraudulentClaim = _
            fraudulentAmount / fraudulentClaims
    End If
    '----------------------------------------------------
    ' TITLE
    '----------------------------------------------------
    wsAnalysis.Range("A1").Value = _
        "INSURANCE CLAIMS ANALYSIS"
    wsAnalysis.Range("A1").Font.Bold = True
    wsAnalysis.Range("A1").Font.Size = 16
    '----------------------------------------------------
    ' PORTFOLIO OVERVIEW
    '----------------------------------------------------
    wsAnalysis.Range("A3").Value = _
        "PORTFOLIO OVERVIEW"
    wsAnalysis.Range("A5").Value = "Total Claims"
    wsAnalysis.Range("B5").Value = totalClaims
    wsAnalysis.Range("A6").Value = "Total Claim Amount"
    wsAnalysis.Range("B6").Value = totalAmount
    wsAnalysis.Range("A7").Value = "Average Claim Amount"
    wsAnalysis.Range("B7").Value = averageAmount
    wsAnalysis.Range("A8").Value = "Median Claim Amount"
    wsAnalysis.Range("B8").Value = medianAmount
    wsAnalysis.Range("A9").Value = "Maximum Claim Amount"
    wsAnalysis.Range("B9").Value = maximumAmount
    wsAnalysis.Range("A10").Value = "Minimum Claim Amount"
    wsAnalysis.Range("B10").Value = minimumAmount
    '----------------------------------------------------
    ' FRAUD OVERVIEW
    '----------------------------------------------------
    wsAnalysis.Range("D3").Value = _
        "FRAUD OVERVIEW"
    wsAnalysis.Range("D5").Value = "Fraudulent Claims"
    wsAnalysis.Range("E5").Value = fraudulentClaims
    wsAnalysis.Range("D6").Value = "Fraud Rate"
    wsAnalysis.Range("E6").Value = fraudRate
    '----------------------------------------------------
    ' CLAIMS BY CLAIM TYPE
    '----------------------------------------------------
    startRow = 13
    wsAnalysis.Cells(startRow, 1).Value = _
        "CLAIMS BY CLAIM TYPE"
    wsAnalysis.Cells(startRow + 1, 1).Value = "Claim Type"
    wsAnalysis.Cells(startRow + 1, 2).Value = "Claims"
    wsAnalysis.Cells(startRow + 1, 3).Value = "Amount"
    outputRow = startRow + 2
    For Each key In dictType.Keys
        wsAnalysis.Cells(outputRow, 1).Value = key
        wsAnalysis.Cells(outputRow, 2).Value = _
            dictType(key)(0)
        wsAnalysis.Cells(outputRow, 3).Value = _
            dictType(key)(1)
        outputRow = outputRow + 1
    Next key
    '----------------------------------------------------
    'CLAIMS BY INCIDENT CAUSE
    '----------------------------------------------------
    wsAnalysis.Cells(startRow, 5).Value = _
        "CLAIMS BY INCIDENT CAUSE"
    wsAnalysis.Cells(startRow + 1, 5).Value = "Cause"
    wsAnalysis.Cells(startRow + 1, 6).Value = "Claims"
    wsAnalysis.Cells(startRow + 1, 7).Value = "Amount"
    outputRow = startRow + 2
    For Each key In dictCause.Keys
        wsAnalysis.Cells(outputRow, 5).Value = key
        wsAnalysis.Cells(outputRow, 6).Value = _
            dictCause(key)(0)
        wsAnalysis.Cells(outputRow, 7).Value = _
            dictCause(key)(1)
        outputRow = outputRow + 1
    Next key
    '----------------------------------------------------
    ' CLAIMS BY AREA
    '----------------------------------------------------
    startRow = 25
    wsAnalysis.Cells(startRow, 1).Value = _
        "CLAIMS BY AREA"
    wsAnalysis.Cells(startRow + 1, 1).Value = "Area"
    wsAnalysis.Cells(startRow + 1, 2).Value = "Claims"
    wsAnalysis.Cells(startRow + 1, 3).Value = "Amount"
    outputRow = startRow + 2
    For Each key In dictArea.Keys
        wsAnalysis.Cells(outputRow, 1).Value = key
        wsAnalysis.Cells(outputRow, 2).Value = _
            dictArea(key)(0)
        wsAnalysis.Cells(outputRow, 3).Value = _
            dictArea(key)(1)
        outputRow = outputRow + 1
    Next key
    '----------------------------------------------------
    ' CLAIMS BY POLICE REPORT
    '----------------------------------------------------
    startRow = 25
    wsAnalysis.Cells(startRow, 5).Value = _
        "CLAIMS BY POLICE REPORT"
    wsAnalysis.Cells(startRow + 1, 5).Value = "Report"
    wsAnalysis.Cells(startRow + 1, 6).Value = "Claims"
    wsAnalysis.Cells(startRow + 1, 7).Value = "Amount"
    outputRow = startRow + 2
    For Each key In dictPolice.Keys
        wsAnalysis.Cells(outputRow, 5).Value = key
        wsAnalysis.Cells(outputRow, 6).Value = _
            dictPolice(key)(0)
        wsAnalysis.Cells(outputRow, 7).Value = _
            dictPolice(key)(1)
        outputRow = outputRow + 1
    Next key
    '----------------------------------------------------
    ' CLAIMS BY SEVERITY
    '----------------------------------------------------
    startRow = 33
    wsAnalysis.Cells(startRow, 1).Value = _
        "CLAIMS BY SEVERITY"
    wsAnalysis.Cells(startRow + 1, 1).Value = "Severity"
    wsAnalysis.Cells(startRow + 1, 2).Value = "Claims"
    wsAnalysis.Cells(startRow + 1, 3).Value = "Amount"
    outputRow = startRow + 2
    For Each key In dictSeverity.Keys
        wsAnalysis.Cells(outputRow, 1).Value = key
        wsAnalysis.Cells(outputRow, 2).Value = _
            dictSeverity(key)(0)
        wsAnalysis.Cells(outputRow, 3).Value = _
            dictSeverity(key)(1)
        outputRow = outputRow + 1
    Next key
    '----------------------------------------------------
    ' CUSTOMER / POLICY ANALYSIS
    '----------------------------------------------------
    startRow = 33
    wsAnalysis.Cells(startRow, 5).Value = _
        "CUSTOMER / POLICY ANALYSIS"
    wsAnalysis.Cells(startRow + 2, 5).Value = _
        "Total Customers"
    wsAnalysis.Cells(startRow + 2, 6).Value = _
        dictCustomer.Count
    If dictCustomer.Count > 0 Then
        wsAnalysis.Cells(startRow + 3, 5).Value = _
            "Average Claims per Customer"
        wsAnalysis.Cells(startRow + 3, 6).Value = _
            totalClaims / dictCustomer.Count
    Else
        wsAnalysis.Cells(startRow + 3, 5).Value = _
            "Average Claims per Customer"
        wsAnalysis.Cells(startRow + 3, 6).Value = 0
    End If
    multipleClaimCustomers = 0
    For Each key In dictCustomer.Keys
        If dictCustomer(key) > 1 Then
            multipleClaimCustomers = _
                multipleClaimCustomers + 1
        End If
    Next key
    wsAnalysis.Cells(startRow + 4, 5).Value = _
        "Customers with Multiple Claims"
    wsAnalysis.Cells(startRow + 4, 6).Value = _
        multipleClaimCustomers
    '----------------------------------------------------
    ' TIME ANALYSIS
    '----------------------------------------------------
    startRow = 42
    wsAnalysis.Cells(startRow, 1).Value = _
        "TIME ANALYSIS"
    wsAnalysis.Cells(startRow + 1, 1).Value = "Year"
    wsAnalysis.Cells(startRow + 1, 2).Value = "Claims"
    wsAnalysis.Cells(startRow + 1, 3).Value = "Total Amount"
    wsAnalysis.Cells(startRow + 1, 4).Value = "Average Amount"
    outputRow = startRow + 2
    For Each key In dictYear.Keys
        wsAnalysis.Cells(outputRow, 1).Value = key
        wsAnalysis.Cells(outputRow, 2).Value = _
            dictYear(key)(0)
        wsAnalysis.Cells(outputRow, 3).Value = _
            dictYear(key)(1)
        If dictYear(key)(0) > 0 Then
            wsAnalysis.Cells(outputRow, 4).Value = _
                dictYear(key)(1) / dictYear(key)(0)
        Else
            wsAnalysis.Cells(outputRow, 4).Value = 0
        End If
        outputRow = outputRow + 1
    Next key
    '----------------------------------------------------
    ' FRAUD ANALYSIS
    '----------------------------------------------------
    startRow = 50
    wsAnalysis.Cells(startRow, 1).Value = _
        "FRAUD ANALYSIS"
    wsAnalysis.Cells(startRow + 2, 1).Value = _
        "Fraudulent Claims"
    wsAnalysis.Cells(startRow + 2, 2).Value = _
        fraudulentClaims
    wsAnalysis.Cells(startRow + 3, 1).Value = _
        "Fraud Rate"
    wsAnalysis.Cells(startRow + 3, 2).Value = _
        fraudRate
    wsAnalysis.Cells(startRow + 4, 1).Value = _
        "Fraudulent Claim Amount"
    wsAnalysis.Cells(startRow + 4, 2).Value = _
        fraudulentAmount
    wsAnalysis.Cells(startRow + 5, 1).Value = _
        "Average Fraudulent Claim"
    wsAnalysis.Cells(startRow + 5, 2).Value = _
        averageFraudulentClaim
    '----------------------------------------------------
    ' FRAUD VS NON-FRAUD
    '----------------------------------------------------
    wsAnalysis.Cells(startRow + 7, 1).Value = _
        "Fraudulent vs Non-Fraudulent"
    wsAnalysis.Cells(startRow + 8, 1).Value = "Status"
    wsAnalysis.Cells(startRow + 8, 2).Value = "Claims"
    wsAnalysis.Cells(startRow + 8, 3).Value = "Amount"
    wsAnalysis.Cells(startRow + 9, 1).Value = _
        "Fraudulent"
    wsAnalysis.Cells(startRow + 9, 2).Value = _
        fraudulentClaims
    wsAnalysis.Cells(startRow + 9, 3).Value = _
        fraudulentAmount
    wsAnalysis.Cells(startRow + 10, 1).Value = _
        "Non-Fraudulent"
    wsAnalysis.Cells(startRow + 10, 2).Value = _
        nonFraudulentClaims
    wsAnalysis.Cells(startRow + 10, 3).Value = _
        nonFraudulentAmount
    '====================================================
    ' FORMATTING
    '====================================================
    wsAnalysis.Range("A3:G3").Font.Bold = True
    wsAnalysis.Range("A13:G14").Font.Bold = True
    wsAnalysis.Range("A25:G26").Font.Bold = True
    wsAnalysis.Range("A33:G34").Font.Bold = True
    wsAnalysis.Range("A42:D43").Font.Bold = True
    wsAnalysis.Range("A50:C50").Font.Bold = True
    wsAnalysis.Range("A1:G1").Font.Bold = True
    wsAnalysis.Range("B6:B10").NumberFormat = _
        "$#,##0.00"
    wsAnalysis.Range("C15:C100").NumberFormat = _
        "$#,##0.00"
    wsAnalysis.Range("G15:G100").NumberFormat = _
        "$#,##0.00"
    wsAnalysis.Range("C27:C100").NumberFormat = _
        "$#,##0.00"
    wsAnalysis.Range("G27:G100").NumberFormat = _
        "$#,##0.00"
    wsAnalysis.Range("C35:C100").NumberFormat = _
        "$#,##0.00"
    wsAnalysis.Range("C44:D100").NumberFormat = _
        "$#,##0.00"
    wsAnalysis.Range("B54:B60").NumberFormat = _
        "$#,##0.00"
    wsAnalysis.Range("E6").NumberFormat = "0.00%"
    wsAnalysis.Range("B53").NumberFormat = "0.00%"
    wsAnalysis.Columns("A:G").AutoFit
    
    MsgBox "Claims analysis updated successfully.", _
           vbInformation
End Sub

