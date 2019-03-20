Sub StockCompilerF():


Dim ws As Worksheet
Dim lrRaw As Long
Dim sTotal As Integer
Dim sWorked As Integer
Dim rWorked As Long
Dim lWorked As Long
Dim startTime As Double
Dim minElapsed As String

startTime = Timer
sWorked = 0
sTotal = ThisWorkbook.Sheets.Count


For Each ws In ActiveWorkbook.Worksheets
  
    sWorked = sWorked + 1
  
    'Set all headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'Find the last non-blank cell in column A(1)
    lrRaw = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    lWorked = lWorked + (lrRaw - 1)
    
    'initialize first ticker name
    ws.Range("I2").Value = ws.Range("A2").Value
    'initialze first ticker open value
    ws.Range("J2").Value = ws.Range("C2").Value
    
  
'**column name cheatsheet**
'1: ticker
'2: Date
'3: open
'4: high
'5: low
'6: Close
'7: vol
'8: null
'9: ticker
'10: yearly Change
'11: Percent Change
'12: total stock volume
        
        
        
'***COMPILER***

    Dim t As Integer
    t = 2
    'for each column A
    For r = 2 To lrRaw
       

       
       'if raw ticker matches total ticker
        If ws.Cells(t, 9).Value = ws.Cells(r, 1).Value Then
            'add daily volume to total volume
            ws.Cells(t, 12).Value = ws.Cells(t, 12).Value + ws.Cells(r, 7).Value
            
            Application.StatusBar = "Processing... " & Round((r / lrRaw * 100), 0) & "% Sheet " & sWorked & " of " & sTotal & " | " & (rWorked) & " stocks compiled."
            
            'if working iteration is last on sheet
            If r = lrRaw Then
                'use close value from this row to calculate delta
                
                'Div/0 Error handler
                If ws.Cells(t, 10).Value = 0 Then
                    ws.Cells(t, 10).Value = 0
                    ws.Cells(t, 11).Value = 0
                    ws.Cells(t, 10).Interior.ColorIndex = 6
                  
                
                    Else
                
                
                    'percent change
                    ws.Cells(t, 11).Value = (ws.Cells(r, 6).Value - ws.Cells(t, 10).Value) / ws.Cells(t, 10).Value
                    'annual change + conditional formatting
                    ws.Cells(t, 10).Value = ws.Cells(r, 6).Value - ws.Cells(t, 10).Value
                    ws.Cells(t, 10).NumberFormat = "0.00"
                    ws.Cells(t, 11).NumberFormat = "0.00%"
                
                        If ws.Cells(t, 10).Value > 0 Then
                            ws.Cells(t, 10).Interior.ColorIndex = 4
                            Else
                            ws.Cells(t, 10).Interior.ColorIndex = 3
                        End If
                End If

                If ws.Cells(t, 11).Value > ws.Range("Q2").Value Then
                    ws.Range("P2").Value = ws.Cells(t, 9).Value
                    ws.Range("Q2").Value = ws.Cells(t, 11).Value

                    ElseIf ws.Cells(t, 11).Value < ws.Range("Q3").Value Then
                    ws.Range("P3").Value = ws.Cells(t, 9).Value
                    ws.Range("Q3").Value = ws.Cells(t, 11).Value
                End If

                If ws.Cells(t, 12).Value > ws.Range("Q4").Value Then
                    ws.Range("P3").Value = ws.Cells(t, 9).Value
                    ws.Range("Q4").Value = ws.Cells(t, 12).Value
                End If

                rWorked = rWorked + 1

            End If
            
        Else
            'calculate yearly change of previous ticker value using previously stored open value
            
            
            'Div/0 Error handler
            If ws.Cells(t, 10).Value = 0 Then
                    ws.Cells(t, 10).Value = 0
                    ws.Cells(t, 11).Value = 0
                    ws.Cells(t, 10).Interior.ColorIndex = 6
                    
                Else
                'percent change
                ws.Cells(t, 11).Value = (ws.Cells(r - 1, 6).Value - ws.Cells(t, 10).Value) / ws.Cells(t, 10).Value
                'annual change
                ws.Cells(t, 10).Value = ws.Cells(r - 1, 6).Value - ws.Cells(t, 10).Value
                ws.Cells(t, 10).NumberFormat = "0.00"
                ws.Cells(t, 11).NumberFormat = "0.00%"
                    If ws.Cells(t, 10).Value > 0 Then
                        ws.Cells(t, 10).Interior.ColorIndex = 4
                        Else
                        ws.Cells(t, 10).Interiocondar.ColorIndex = 3
                    End If
            End If
                
            If ws.Cells(t, 11).Value > ws.Range("Q2").Value Then
                    ws.Range("P2").Value = ws.Cells(t, 9).Value
                    ws.Range("Q2").Value = ws.Cells(t, 11).Value

                ElseIf ws.Cells(t, 11).Value < ws.Range("Q3").Value Then
                    ws.Range("P3").Value = ws.Cells(t, 9).Value
                    ws.Range("Q3").Value = ws.Cells(t, 11).Value
                End If

            If ws.Cells(t, 12).Value > ws.Range("Q4").Value Then
                ws.Range("P4").Value = ws.Cells(t, 9).Value
                ws.Range("Q4").Value = ws.Cells(t, 12).Value
                End If
                
            '+1 to total row count
             t = t + 1
            
            'add new stock ticker ID to list
            ws.Cells(t, 9).Value = ws.Cells(r, 1).Value
            'add volume value to new row
            ws.Cells(t, 12).Value = ws.Cells(t, 12).Value + ws.Cells(r, 7).Value
            'store next opening value to new row
            ws.Cells(t, 10).Value = ws.Cells(r, 3)
                  
            rWorked = rWorked + 1

        End If
            
    Next r

Next ws

Application.StatusBar = ""
minElapsed = Format((Timer - startTime) / 86400, "hh:mm:ss")
MsgBox ("Processing Complete. " & vbNewLine & lWorked & " records compiled." & vbNewLine & "Time Elapsed: " & minElapsed)
 


End Sub