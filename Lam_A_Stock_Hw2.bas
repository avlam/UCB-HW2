Attribute VB_Name = "Module1"
Sub stocks()
'Output Main summary in columns I through L (columns 9 through 12)
'Output for overall by page in columns O through Q (columns 14 through 16)

'Loop through all ticker entries
'Sum total volume per unique ticker symbol
'Track opening value (column 3, minimum date)
'Track and compare to closing value (column 6, max date)
'keep track of years...?
'Summarize with ticker symbol

'Track Current Greatest volume, %increase and %decrease

'Assumed: In a given worksheet, no more than a year's data will be provided for a single symbol
'Assumed: Tickets are ordered; once the symbol changes it will not appear again in the worksheet




For Each ws In Worksheets
'Print Headers for Output
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Volume"

ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"


'Initialize counters, and variables
'Current Ticker Symbol
Dim thisTicker As String
Dim thisVol As Double
thisVol = 0 'initialize volume for first symbol
Dim thisOpen As Double
Dim thisClose As Double
Dim firstTicker As Boolean


'Other Outputs
Dim tChange As Double 'Total Change (Per Symbol)
Dim pChange As Double 'Percent Change (Per Symbol)


'RowIndex of Summary Output
Dim SummaryOut As Double


'Last Row of the Page
Dim LastRow As Double
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'create and initialize Candidate holders for "Greatest" outputs
Dim cand_greatest_inc_val As Double
cand_greatest_inc_val = 0
Dim cand_greatest_inc_sym As String
Dim cand_greatest_dec_val As Double
cand_greatest_dec_val = 0
Dim cand_greatest_dec_sym As String
Dim cand_greatest_vol_val As Double
cand_greatest_vol_val = 0
Dim cand_greatest_vol_sym As String


'iterate over each row
    For thisRow = 2 To LastRow
'Special Case for first and last table entry
        If thisRow = 2 Then
            SummaryOut = 2 'Start on row immediately below header
            firstTicker = True
        End If
        
'Track Running Totals for pChange and tChange
        thisTicker = ws.Cells(thisRow, 1).Value
        thisVol = thisVol + ws.Cells(thisRow, 7).Value
        If firstTicker Then
            thisOpen = ws.Cells(thisRow, 3).Value
            firstTicker = False
        End If
        
'Check if Next row's ticker or if it's the last row

        If ws.Cells(thisRow + 1, 1) <> thisTicker Or thisRow = LastRow Then
            thisClose = ws.Cells(thisRow, 6).Value
            tChange = thisClose - thisOpen
            If thisOpen <> 0 Then
                pChange = tChange / thisOpen
            End If
            
            'Print Summary for current Ticker if it changes
            ws.Cells(SummaryOut, 9).Value = thisTicker
            ws.Cells(SummaryOut, 10).Value = tChange
            
            'Format for +/-
            If tChange > 0 Then
                ws.Cells(SummaryOut, 10).Interior.ColorIndex = 4
            ElseIf tChange < 0 Then
                ws.Cells(SummaryOut, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(SummaryOut, 10).Interior.ColorIndex = 0
            End If

            'Include Exception for zero opening value
            If thisOpen <> 0 Then
                ws.Cells(SummaryOut, 11).Value = pChange
            Else
                ws.Cells(SummaryOut, 11).Value = "N/A"
                MsgBox (ws.Cells(thisRow, 1).Value + " Has 0 Opening value for " + ws.Name)
            End If
            
            ws.Cells(SummaryOut, 11).NumberFormat = "0.00%"
            ws.Cells(SummaryOut, 12).Value = thisVol
            'Check for "Greatests"
            If thisVol > cand_greatest_vol_val Then
                cand_greatest_vol_val = thisVol
                cand_greatest_vol_sym = thisTicker
            End If
            If pChange > cand_greatest_inc_val Then
                cand_greatest_inc_val = pChange
                cand_greatest_inc_sym = thisTicker
            ElseIf pChange < cand_greatest_dec_val Then
                cand_greatest_dec_val = pChange
                cand_greatest_dec_sym = thisTicker
            End If
            'Reset for next ticker symbol
            thisVol = 0
            thisOpen = 0
            thisClose = 0
            SummaryOut = SummaryOut + 1
            firstTicker = True
        End If
    Next thisRow
'Print "Greatests"
ws.Cells(2, 15).Value = cand_greatest_inc_sym
ws.Cells(3, 15).Value = cand_greatest_dec_sym
ws.Cells(4, 15).Value = cand_greatest_vol_sym
ws.Cells(2, 16).Value = cand_greatest_inc_val
ws.Cells(3, 16).Value = cand_greatest_dec_val
ws.Cells(4, 16).Value = cand_greatest_vol_val
ws.Range("P2:P3").NumberFormat = "0.00%"
Next ws
End Sub
