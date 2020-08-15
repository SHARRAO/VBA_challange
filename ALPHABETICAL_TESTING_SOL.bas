Attribute VB_Name = "Module1"
Sub Test_data():

'Part 1
'To get unique ticker values for a single sheet
    'Declaring variables for tickervalue
Dim n As Integer
Dim x As Long
Dim volume As Double 'declaring variable for total volume of a stock PART5

    'Declarin variables for yearly change
    
Dim OP As Double 'OP is the opening price at the begining of the year
Dim CP As Double 'CP is closing price at the end of the year
    

Dim Grtper_inc As Double 'declaring variables for gratest %
Dim lowper_inc As Double 'declaring variables for lowest %
Dim Gratest_tol_vol As Double 'declaring variable for gratest total volume
Dim grtper_tic As String 'declaring gratest % ticker value
Dim lowper_tic As String 'declaring lowest % ticker value
Dim grtvol_tic As String 'declaring gratest ticker value

For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    
    ws.Range("k1").Value = "open price if the ticker begining year" 'added these coloum for ease of reading
    ws.Range("l1").Value = "close price of the ticker year ending"  'added these coloum for ease of reading
    ws.Range("M1").Value = "Percentage Change"
    ws.Range("N1").Value = "Total stock volume of the stock "
    ws.Range("R1").Value = "Ticker"
    ws.Range("S1").Value = "Value"
    ws.Range("q2").Value = "Gratest % increase"
    ws.Range("q3").Value = "Lowest % decrease"
    ws.Range("q4").Value = "Gratest Total Volume"
        
    n = 1
    Grtper_inc = 0
    lowper_inc = 0
    Gratest_tol_vol = 0
    grtper_tic = ""
    lowper_tic = ""
    grtvol_tic = ""
'last used row of current worksheet
    x = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To x
    If i = 2 Or ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        n = n + 1
        ws.Cells(n, 9).Value = ws.Cells(i, 1).Value
        OP = ws.Cells(i, 3).Value
        ws.Cells(n, 11) = OP
        
        volume = ws.Cells(i, 7).Value          'Part 5
                                            'Total stock volume of the stock
                                            'summation of volume for each day (from each cell) of a similar ticker
                                            
    End If
    volume = volume + ws.Cells(i, 7)
    
    'PART 7
'GRATER total volume
        If ws.Cells(n, 14).Value > Gratest_tol_vol Then
            Gratest_tol_vol = ws.Cells(n, 14).Value
            grtvol_tic = ws.Cells(n, 9).Value
        End If


    
'Part 2
'Yearly change based on opening price at a begining of a year
'To closing price at end of the year

    If i = x Or ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        CP = ws.Cells(i, 6).Value
        ws.Cells(n, 12) = CP
        ws.Cells(n, 10).Value = CP - OP
        
        'Part 4
 'Percentage Change
   If OP > 0 Then
   ws.Cells(n, 13).Value = Round((CP - OP) / OP, 2)
   ws.Cells(n, 13).NumberFormat = "0.00%"                   'Percent type
    Else: ws.Cells(n, 13).Value = 0
    End If
    ws.Cells(n, 14).Value = volume
        
'PART 5
'GRATER % INCREASE CODE
        
        If ws.Cells(n, 13).Value > Grtper_inc Then
           Grtper_inc = ws.Cells(n, 13).Value
           grtper_tic = ws.Cells(n, 9).Value
           ws.Cells(n, 19).NumberFormat = "0.00%"
    
        End If
     
   
'PART 6
'LOWEST % DECREASE CODE
    
    If ws.Cells(n, 13).Value < lowper_inc Then
           lowper_inc = ws.Cells(n, 13).Value
           lowper_tic = ws.Cells(n, 9).Value
           
           
    End If
     
    
    
 
'Part 3
'Colour coding based on values
 
        If OP > CP Then
            ws.Range(Cells(n, 10).Address).Interior.ColorIndex = 3  '-VE VALUES
            
    Else:
            ws.Range(Cells(n, 10).Address).Interior.ColorIndex = 4 '+VE VALUES
            
        End If
    End If

Next i


ws.Cells(2, 18).Value = grtper_tic
ws.Cells(3, 18).Value = lowper_tic
ws.Cells(4, 18).Value = grtvol_tic
ws.Cells(2, 19).Value = Grtper_inc
ws.Cells(3, 19).Value = lowper_inc
ws.Cells(4, 19).Value = Gratest_tol_vol
Next ws

End Sub


