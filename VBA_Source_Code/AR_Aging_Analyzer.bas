' ===================================================================================
' MODULE:  AR_Aging_Analyzer
' AUTHOR:  Chaitanya Yarlagadda
' PURPOSE: Ages open AR invoices, flags anything 90+ days as write-off risk, and
'          writes a summary a finance manager can read without digging.
'
' SHEET LAYOUT
'   A Invoice_ID | B Customer_Name | C Payer_Type | D Invoice_Date | E Balance_Due
'   The macro writes F and G. The summary goes in I, J, K.
'
' NOTE: First VBA project. I used AI to get started since I had not written VBA
'       before, then went through it line by line and fixed what broke.
' ===================================================================================

Option Explicit

Private Const RISK_THRESHOLD_DAYS As Long = 90
Private Const DATA_SHEET As String = "Sheet1"


Sub RunAgingReport()

    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, age As Long
    Dim balance As Double
    Dim totalAR As Double, criticalAR As Double, activeAR As Double
    Dim pctOver90 As Double

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    ' I name the sheet instead of using ActiveSheet. The button sits on the
    ' Dashboard, so ActiveSheet would be the wrong sheet when it runs.
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)

    ' Counting from column B because customer name is always filled in.
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data found on " & DATA_SHEET & ".", vbExclamation
        GoTo CleanExit
    End If

    ws.Range("F1").Value = "Days_Outstanding"
    ws.Range("G1").Value = "AR_Status_Alert"
    ws.Range("F1:G1").Font.Bold = True

    ' Clearing the whole column, not a fixed range like F2:G100. A fixed number
    ' works until the data goes past it - same trap I hit with the SUMIF range.
    ws.Range("F2:G" & ws.Rows.Count).ClearContents
    ws.Range("F2:G" & ws.Rows.Count).Interior.ColorIndex = xlNone

    For i = 2 To lastRow

        If Len(Trim(ws.Cells(i, 2).Value)) > 0 Then

            ' This is age from the invoice date, so days outstanding, not days
            ' past due. Past due would need payment terms and I do not have those.
            ' IsDate first, so one bad date does not stop the whole run.
            If IsDate(ws.Cells(i, 4).Value) Then
                age = Date - CDate(ws.Cells(i, 4).Value)
            Else
                age = 0
            End If

            ' Format has to be set before the value goes in. Otherwise the cell
            ' picks up date formatting from the subtraction and shows 03/02/1900
            ' instead of 62. Took me a while to figure out.
            ws.Cells(i, 6).NumberFormat = "0"
            ws.Cells(i, 6).Value = age

            ' Stripping $ and commas in case the balance comes in as text from
            ' an export. Val on its own returns 0 for anything with a $ in front.
            balance = Val(Replace(Replace(Replace( _
                          ws.Cells(i, 5).Value, "$", ""), ",", ""), " ", ""))

            If balance <= 0 Then
                ws.Cells(i, 7).Value = "Settled"
                ws.Cells(i, 7).Interior.ColorIndex = xlNone

            ElseIf age >= RISK_THRESHOLD_DAYS Then
                ws.Cells(i, 7).Value = "CRITICAL: Write-off Risk"
                ws.Cells(i, 7).Interior.Color = RGB(255, 204, 204)
                criticalAR = criticalAR + balance

            Else
                ws.Cells(i, 7).Value = "Active Collection"
                ws.Cells(i, 7).Interior.Color = RGB(204, 255, 204)
                activeAR = activeAR + balance
            End If

            totalAR = totalAR + balance

        End If
    Next i

    ' --- summary -------------------------------------------------------------
    ' Adding the totals up inside the loop instead of writing SUMIF formulas
    ' to the sheet. I had SUMIF working, but the range was fixed at row 5 from
    ' when I built it on 5 rows of test data. After I pasted in 15, exposure
    ' showed $36,350 instead of $67,100 and nothing looked wrong on screen.
    ' Widening the range fixed it, but the range can go stale again the next
    ' time the data grows. Adding as the loop goes cannot skip a row it has
    ' already been through.

    ws.Range("I1").Value = "Metric Description"
    ws.Range("J1").Value = "Value"
    ws.Range("K1").Value = "Context/Target"
    ws.Range("I1:K1").Font.Bold = True

    ws.Range("I2").Value = "Total Outstanding AR"
    ws.Range("J2").Value = totalAR
    ws.Range("K2").Value = "Gross Ledger Balance"

    ws.Range("I3").Value = "Critical Write-off Exposure"
    ws.Range("J3").Value = criticalAR
    ws.Range("K3").Value = "Breached 90+ Day Target"

    ws.Range("I4").Value = "Active Collections Pipeline"
    ws.Range("J4").Value = activeAR
    ws.Range("K4").Value = "Within 90 days"

    ' This is the one an AR manager actually watches each month. If it climbs,
    ' collections are falling behind billing.
    If totalAR > 0 Then pctOver90 = criticalAR / totalAR
    ws.Range("I5").Value = "% of AR over 90 days"
    ws.Range("J5").Value = pctOver90
    ws.Range("K5").Value = "Under 25%"

    ws.Range("J2:J4").NumberFormat = "$#,##0.00"
    ws.Range("J5").NumberFormat = "0%"

    ' --- tie-out check --------------------------------------------------------
    ' Critical plus Active has to come back to Total AR. If it does not, a row
    ' went into neither one or got counted twice. This is the check that would
    ' have caught my SUMIF problem straight away, so I put it in.
    ws.Range("I7").Value = "Check: components tie to total"
    If Abs((criticalAR + activeAR) - totalAR) < 0.01 Then
        ws.Range("J7").Value = "OK"
        ws.Range("J7").Interior.Color = RGB(204, 255, 204)
    Else
        ws.Range("J7").Value = "MISMATCH: " & _
            Format((criticalAR + activeAR) - totalAR, "$#,##0.00")
        ws.Range("J7").Interior.Color = RGB(255, 204, 204)
    End If

    ws.Columns("A:K").AutoFit

    ' The macro stamps the date instead of =NOW() on the sheet. NOW recalculates
    ' any time anything in the workbook changes, so it would show today's date
    ' sitting on top of last week's numbers. Writing it here means the date only
    ' moves when the aging actually runs.
    With ThisWorkbook.Worksheets("Dashboard").Range("B3")
        .Value = Date
        .NumberFormat = "m/d/yy"
    End With

CleanExit:
    Application.ScreenUpdating = True
    MsgBox "Aging report complete." & vbCrLf & vbCrLf & _
           "Total AR: " & Format(totalAR, "$#,##0.00") & vbCrLf & _
           "At risk (90+ days): " & Format(criticalAR, "$#,##0.00"), _
           vbInformation, "AR Aging Analyzer"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical, "AR Aging Analyzer"
End Sub

