Sub Populate_All_Worksheets_Reattempt()

Dim ws As Worksheet

'loop through each worksheet, from first to last worksheet
For Each ws In ThisWorkbook.Worksheets

'Run the following steps for each worksheet
'NOTE. Tutor (Marc Calache) advised to add "ws." in front of all Cells and Range objects, to enable looping through all procedures.

'---------------------------------------------------------------------------------
'STEP 1. Find unique ticker names in column A and paste into column I

Dim tickerList As Range

    'count every row in column A that contains a ticker name
    Set tickerList = ws.Range("A2:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)

    'get unique values from column A and paste in column I from row 2
    tickerList.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ws.Range("I1"), Unique:=True

'Note. last action duplicates the first ticker in the list.
'The next step will override the duplicated ticker by placing a header called "Ticker" in cell "I1"

'---------------------------------------------------------------------------------
'STEP 2. Add headers to table in columns I to L, and for bonus table O to Q
    
    With ws
    
        .Range("I1").Value = "Ticker"
        .Range("J1").Value = "Yearly Change"
        .Range("K1").Value = "Percent Change"
        .Range("L1").Value = "Total Stock Volume"
       
        .Range("O2").Value = "Greatest % Increase"
        .Range("O3").Value = "Greatest % Decrease"
        .Range("O4").Value = "Greatest Total Volume"
        .Range("P1").Value = "Ticker"
        .Range("Q1").Value = "Value"
        
    End With
          
'---------------------------------------------------------------------------------
'STEP 3. Calculate Yearly Change in Stock Value for each unique ticker
'2nd attempt changed begticker and endticker from integer to double

Dim begticker As Double
Dim endticker As Double
Dim k As Integer
Dim uniquetickers As Integer

    'count every row in column I that contains a ticker name
    uniquetickers = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
        For k = 2 To uniquetickers

        'find the first row the ticker name appears in column A from top to bottom
        '2nd attempt added "LookIn:=xlValues, LookAt:=xlWhole" to look in the cell values and search for an exact match of the entire call contents
        begticker = ws.Range("A:A").Find(What:=ws.Cells(k, 9), After:=ws.Range("A1"), LookIn:=xlValues, LookAt:=xlWhole).Row
        
        'find the last row the ticker name appears in column A by searching the list bottom to top
        '2nd attempt added "LookIn:=xlValues, LookAt:=xlWhole" to look in the cell values and search for an exact match of the entire call contents
        endticker = ws.Range("A:A").Find(What:=ws.Cells(k, 9), After:=ws.Range("A1"), LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious).Row

        'calculate Yearly Change
        ws.Cells(k, 10) = ws.Cells(endticker, 6) - ws.Cells(begticker, 3)

        'Add conditional formatting to cells in column J, red for negative and green for positive
        If ws.Cells(k, 10) < 0 Then
            ws.Cells(k, 10).Interior.Color = RGB(255, 0, 0)
        Else
            ws.Cells(k, 10).Interior.Color = RGB(0, 255, 0)
        End If

        'calculate Percent Change
        ws.Cells(k, 11) = ((ws.Cells(endticker, 6) - ws.Cells(begticker, 3)) / ws.Cells(begticker, 3))

        'Change format of cells in column K to percentage with 2 decimal places
         ws.Cells(k, 11).NumberFormat = "0.00%"

        Next k

'---------------------------------------------------------------------------------
'STEP 4. Sum up Total Stock Volume from column G for each unique tickers in column I

Dim j As Integer

    'count every row in column I that contains a ticker name
    'uniquetickers = Cells(Rows.Count, "I").End(xlUp).Row

    For j = 2 To uniquetickers
        ws.Cells(j, 12) = WorksheetFunction.SumIf(ws.Range("A:A"), ws.Cells(j, 9), ws.Range("G:G"))
    Next j
    
'---------------------------------------------------------------------------------
'STEP 5. Populate bonus table to the right from column O to Q

'uniquetickers = Cells(Rows.Count, "I").End(xlUp).Row

'Find Ticker with Greatest % increase (Ticker from column I, % increase from column K)
Dim increase_number As String
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & uniquetickers)), ws.Range("K2:K" & uniquetickers), 0)
        ws.Range("P2") = ws.Cells(increase_number + 1, 9)

    'Find corresponding Greatest % increase in percentage format
    Dim greatestinc As Double
        ws.Range("Q2") = WorksheetFunction.Max(ws.Range("K:K"))
        ws.Range("Q2").NumberFormat = "0.00%"
    
'Ticker with greatest % Decrease (Ticker from column I, % decrease from column K)
Dim decrease_number As String
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & uniquetickers)), ws.Range("K2:K" & uniquetickers), 0)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9)

    'Find corresponding Greatest % decrease in percentage format
    Dim greatestdec As Double
        ws.Range("Q3") = WorksheetFunction.Min(ws.Range("K:K"))
        ws.Range("Q3").NumberFormat = "0.00%"

'Ticker with Greatest Total Volume (Ticker from column I, % decrease from column L)
Dim increase_number2 As String
    increase_number2 = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & uniquetickers)), ws.Range("L2:L" & uniquetickers), 0)
        ws.Range("P4") = ws.Cells(increase_number2 + 1, 9)

    'Find corresponding Greatest Total Volume in scientific format
    Dim greatestvol As Double
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L:L"))
        ws.Range("Q4").NumberFormat = "0.00E+00"

'Adjust width of all columns to see data properly
    ws.Columns("I:I").ColumnWidth = 6.14
    ws.Columns("J:J").ColumnWidth = 12.86
    ws.Columns("K:K").ColumnWidth = 14.29
    ws.Columns("L:L").ColumnWidth = 17.57
    ws.Columns("M:N").ColumnWidth = 8.43
    ws.Columns("O:O").ColumnWidth = 20.43
    ws.Columns("P:P").ColumnWidth = 6.14
    ws.Columns("Q:Q").ColumnWidth = 9
'---------------------------------------------------------------------------------

        
    Next ws
End Sub