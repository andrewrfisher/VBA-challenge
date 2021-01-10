Attribute VB_Name = "Module1"
Sub WallstreetLoops():


    'set Current WS as worksheet object variable
    Dim Current_WS As Worksheet
    
    'figure out what this does
    Dim Need_Summary_Table_Header As Boolean
    Dim COMMAND_SPREADSHEET As Boolean
    
    'set header and hard part
    Need_Summary_Table_Header = False
    COMMAND_SPREADSHEET = True
    
    'loop through all worksheets in workbook
    For Each Current_WS In Worksheets
    
        'set initial variable for ticker name
        Dim Ticker_Name As String
        Ticker_Name = " "
        
        'set initial varialble for holding the total per ticker name
        Dim Total_Ticker_Volume As Double
        Total_Ticker_Volume = 0
        
        'set new variables
        'open price
            Dim Open_Price As Double
            Open_Price = 0
            
        'close price
            Dim Close_Price As Double
            Close_Price = 0
            
        'change in price
            Dim Delta_Price As Double
            Delta_Price = 0
            
        'percentage change for price
            Dim Delta_Percent As Double
            Delta_Percent = 0
            
        'max ticker
            Dim Max_Ticker_Name As String
            Max_Ticker_Name = " "
            
        'min ticker
            Dim Min_Ticker_Name As String
            Min_Ticker_Name = " "
            
        'max percent
            Dim Max_Percent As Double
            Max_Percent = 0
            
        'min percent
            Dim Min_Percent As Double
            Min_Percent = 0
            
        'max volume
            Dim Max_Volume_Ticker As String
            Max_Volume_Ticker = " "
            Dim Max_Volume As Double
            Max_Volume = 0
            
        'keep track of location for each ticker name
        'in the summary table for the current wb
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        'set initial row count
        Dim lastrow As Long
        
        Dim i As Long
        
        lastrow = Current_WS.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        If Need_Summary_Table_Header Then
       
            'set titles for columns
            Current_WS.Range("I1").Value = "Ticker"
            Current_WS.Range("J1").Value = "Yearly Change"
            Current_WS.Range("K1").Value = "Percent Change"
            Current_WS.Range("L1").Value = "Total Stock Volume"
            
            'set titles for additional summary on the right in current ws
            Current_WS.Range("O2").Value = "Greatest % Increase"
            Current_WS.Range("O3").Value = "Greatest % Decrease"
            Current_WS.Range("O4").Value = "Greatest Total Volume"
            Current_WS.Range("P1").Value = "Ticker"
            Current_WS.Range("Q1").Value = "Value"
            
        Else
            
            Need_Summary_Table_Header = True
        
        End If
        
        'set initial value of Open Price for the first Ticker
        Open_Price = Current_WS.Cells(2, 3).Value
        
        'loop from begining of ws until lastrow
        For i = 2 To lastrow
        
        'check if in same ticker, if not write results to summary table
            If Current_WS.Cells(i + 1, 1).Value <> Current_WS.Cells(i, 1).Value Then
        
                'set ticker name and insert the data into designated cell above
                Ticker_Name = Current_WS.Cells(i, 1).Value
            
                'now calculate price change and percent change
                Close_Price = Current_WS.Cells(i, 6).Value
                Delta_Price = Close_Price - Open_Price
            
                'fixing dividing by 0 issue
                If Open_Price <> 0 Then
                Delta_Percent = (Delta_Price / Open_Price) * 100
            
                End If
            
                'adding equation for total ticker volume
                'total ticker volume was set to 0 above
                Total_Ticker_Volume = Total_Ticker_Volume + Current_WS.Cells(i, 7).Value
            
                'add ticker name to  Summary table in column I
                Current_WS.Range("I" & Summary_Table_Row).Value = Ticker_Name
                Current_WS.Range("J" & Summary_Table_Row).Value = Delta_Price
        
                'yearly change and delta price colors green and red
                If (Delta_Price > 0) Then
                Current_WS.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            
                ElseIf (Delta_Price <= 0) Then
                Current_WS.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
                End If
            
                'print the name in the summary table
                Current_WS.Range("K" & Summary_Table_Row).Value = (CStr(Delta_Percent) & "%")
                Current_WS.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
            
                'add 1 to row of summary table
                Summary_Table_Row = Summary_Table_Row + 1
    
            
                'reseting delta price and close price
                Delta_Price = 0
                Close_Price = 0
            
                'now we need to grab ticker's next open price by capturing value one below i
                Open_Price = Current_WS.Cells(i + 1, 3).Value
                
                'populate summary table
                If (Delta_Percent > Max_Percent) Then
                    Max_Percent = Delta_Percent
                    Max_Ticker_Name = Ticker_Name
            
                ElseIf (Delta_Percent < Min_Percent) Then
                    Min_Percent = Delta_Percent
                    Min_Ticker_Name = Ticker_Name
                    
                End If
                
                
                If (Total_Ticker_Volume > Max_Volume) Then
                    Max_Volume = Total_Ticker_Volume
                    Max_Volume_Ticker = Ticker_Name
                    
                End If
                
                'reset counters
                Delta_Percent = 0
                Total_Ticker_Volume = 0
                
            'if the cell that follow is still the same name add it to the total volume
            Else
                Total_Ticker_Volume = Total_Ticker_Volume + Current_WS.Cells(i, 7).Value
            
            End If
            
            
        Next i
        
            'check if it is not the first ss and record all new counts to summary table
            If Not COMMAND_SPREADSHEET Then
                
                Current_WS.Range("Q2").Value = (CStr(Max_Percent) & "%")
                Current_WS.Range("Q3").Value = (CStr(Min_Percent) & "%")
                Current_WS.Range("P2").Value = Max_Ticker_Name
                Current_WS.Range("P3").Value = Min_Ticker_Name
                Current_WS.Range("P4").Value = Max_Volume_Ticker
                Current_WS.Range("Q4").Value = Max_Volume
              
            Else
                COMMAND_SPREADSHEET = False
            End If
            
            
    Next Current_WS
        
        
        
        
            
        
        

        



End Sub


