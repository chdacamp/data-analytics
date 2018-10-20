Sub StockTickerReal()




'----------------------Loop Through Worksheets--------------------------------------
Dim ws As Worksheet
Dim wscount As Integer

wscount = ActiveWorkbook.Worksheets.Count

For i = 1 To wscount
    
    Set ws = ActiveWorkbook.Worksheets(i)
    
    ws.Activate


'-------------------------------------------------------------------------------------
    
'------------Major Variables Used Throughout------------------------------------------------

    Dim Ticker As String
    Dim Volume As Double
    Dim VList() As Double
    Dim TickList() As String
    Dim Size As Long
    
    Dim rcount As Long

    '----------------------TICKER LIST ARRAY ------------------------------------------------
    
    
        'get number of rows
        rcount = ws.Range("A1", ws.Range("A1").End(xlDown)).Rows.Count
        
        'collect list of unique tickers
        'logic - check if row value exists in array at all, if so grow array and add to last slot
        
        Size = 0
        
        For r = 2 To rcount
            
            Ticker = Cells(r, 1)
            Volume = Cells(r, 7)
            
        'iterate through list, count matches
        
            Dim Matches As Integer
                
            Matches = 0
            
            For s = 0 To Size
            
                'NOTE - preserve is important! otherwise the array gets cleared each loop
                
                ReDim Preserve TickList(Size) As String
                
                'counts up if a match is found
                If Ticker = TickList(s) Then
                    Matches = Matches + 1
                    End If
                
            Next s
                
            'if no match, increase list by 1 and add to new slot
            If Matches = 0 Then
                Size = Size + 1
                ReDim Preserve TickList(Size) As String
                ReDim VList(Size) As Double
                '^sets size for our volume list for later
                TickList(Size) = Ticker
                
            End If
                
            'reset matches counter or it will keep counting up each loop
            Matches = 0
            
        Next r
    
    '----------------------TICKER VOLUME ARRAY ------------------------------------------------
    
        'Similar loop to create a parallel list summing volume when there is a match in the TickList
        
        For r = 2 To rcount
        
            Ticker = Cells(r, 1)
            Volume = Cells(r, 7)
        
            For s = 1 To Size
            
                If TickList(s) = Ticker Then
                    VList(s) = VList(s) + Volume
                    End If
            Next s
                
        Next r
        
    '----------------------STOCK OPEN/CLOSE ARRAYS  ------------------------------------------------
    'creates parallel array for dates value, opening, and closing value. Loop through, and record
    'the lowest date (opening)/highest date (closing) and corresponding opening/closing price.
    'This method can handle an unsorted data set.
    
    Dim Opening As Double
    Dim Closing As Double
    Dim DateNum As Long
    ReDim DateListO(Size) As Long
    ReDim DateListC(Size) As Long
    ReDim OpeningList(Size) As Double
    ReDim ClosingList(Size) As Double
    
        For r = 2 To rcount
        
            Ticker = Cells(r, 1)
            DateNum = Cells(r, 2)
            Opening = Cells(r, 3)
            Closing = Cells(r, 6)
            
            For s = 1 To Size
                
                'OPENING CONDITONAL
                If TickList(s) = Ticker And (DateListO(s) = 0 Or DateListO(s) > DateNum) Then
                    DateListO(s) = DateNum
                    OpeningList(s) = Opening
                End If
                
                'CLOSING CONDITIONAL
                If TickList(s) = Ticker And (DateListC(s) = 0 Or DateListC(s) < DateNum) Then
                    DateListC(s) = DateNum
                    ClosingList(s) = Closing
                End If
    
            Next s
            
        Next r
    
                      
     '----------------------PRINT INITIAL ARRAYS ------------------------------------------------
        
        'now print array in column I, p=1 rather than p=0 because first slot in array is blank
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Volume"
        Cells(1, 11) = "Opening"
        Cells(1, 12) = "Closing"
        
        For p = 1 To Size
        
            Cells(p + 1, 9) = TickList(p)
            Cells(p + 1, 10) = VList(p)
            Cells(p + 1, 11) = OpeningList(p)
            Cells(p + 1, 12) = ClosingList(p)
            
        Next p
    
    '---------------------Difference/Percent Change From Printed Arrays----------------------------------------------'
    
    Dim Difference As Double
    Dim PercentDiff As Double
    
     'get number of rows
    rcount = ws.Range("I1", ws.Range("I1").End(xlDown)).Rows.Count
    
    For r = 2 To rcount
    
            Opening = Cells(r, 11)
            Closing = Cells(r, 12)
            
            'simple evaluation
            Difference = Closing - Opening
            
            '********************avoid 0 denominator before evaluating (Grrr PLNT!!)
            
            If Opening <> 0 Then
                PercentDiff = Difference / Opening
                Else: PercentDiff = 0
            End If
            
            'print results with formatting
            Cells(r, 13) = Difference
            
                If Difference >= 0 Then
                    Cells(r, 13).Interior.ColorIndex = 4
                    Else: Cells(r, 13).Interior.ColorIndex = 3
                End If
            
                Cells(r, 14) = PercentDiff
                Cells(r, 14).NumberFormat = "#.00%"
        
    Next r
    
            'print headers
            Cells(1, 13) = "Yearly Change"
            Cells(1, 14) = "Percent Change"
            
        
    '---------------------Greatest Perc Inc/Dec/Total Volume------------------------------
    
    'Headers
    
    Cells(1, 17) = "Category"
    Cells(1, 18) = "Ticker"
    Cells(1, 19) = "Result"
    Cells(2, 17) = "Greatest % Increase"
    Cells(3, 17) = "Greatest % Decrease"
    Cells(4, 17) = "Greatest Total Volume"
        
    'Variables
    
        Dim IncreaseTicker As String
        Dim IncreaseValue As Double
        
        Dim DecreaseTicker As String
        Dim DecreaseValue As Double
        
        Dim MaxVolTicker As String
        Dim MaxVolValue As Double
    
    For r = 2 To rcount
    
    'replace variable if greater than increase variable (greatest increase captured)
    'inverse for decrease on elseif
    
        If Cells(r, 14) > IncreaseValue Then
            IncreaseValue = Cells(r, 14)
            IncreaseTicker = Cells(r, 9)
        
            ElseIf Cells(r, 14) < DecreaseValue Then
                DecreaseValue = Cells(r, 14)
                DecreaseTicker = Cells(r, 9)
        End If
        
    'same but for greatest volume only
    
        If Cells(r, 10) > MaxVolValue Then
            MaxVolValue = Cells(r, 10)
            MaxVolTicker = Cells(r, 9)
        End If
        
    Next r
        
    'Print Results
    
    Cells(2, 18) = IncreaseTicker
    Cells(2, 19) = IncreaseValue
    Cells(2, 19).NumberFormat = "#.00%"
    
    Cells(3, 18) = DecreaseTicker
    Cells(3, 19) = DecreaseValue
    Cells(3, 19).NumberFormat = "#.00%"
    
    Cells(4, 18) = MaxVolTicker
    Cells(4, 19) = MaxVolValue

'----------************convert undefined "0" to undefined string-----------------------------------------------------------

        For r = 2 To rcount
    
            Opening = Cells(r, 11)
            
            If Opening = 0 Then
                Cells(r, 14) = "div 0, undefined"
            End If
            
        Next r
            


'----------------------end of code to loop------------------
'----------------------------------------------------------
    
Next i
        

End Sub

