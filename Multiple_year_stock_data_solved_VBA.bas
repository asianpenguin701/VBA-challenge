Attribute VB_Name = "Module1"
Sub calculations1()

    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim ticker As String
    Dim startDate As Date
    Dim endDate As Date
    Dim startOpen As Double
    Dim endClose As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim dict As Object
    Dim output As Long
    Dim key As Variant
    Dim Counter As Long
    Dim tickerMaxPercent As String
    Dim tickerMinPercent As String
    Dim tickerMaxVolume As String
    Dim percentMax As Double
    Dim percentMin As Double
    Dim maxVolume As Double
    Dim ws As Worksheet

    ' Initialize/reset variables for each worksheet
    For Each ws In Worksheets

        output = 2
        percentMax = -1
        percentMin = 1
        maxVolume = 0

        ' Find the last row with data in column A
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Create a dictionary object to store results
        Set dict = CreateObject("Scripting.Dictionary")

        ' Loop through each row
        For i = 2 To lastRow ' For loop from row 2 to last row
            ticker = ws.Cells(i, 1).Value ' Ticker
            startDate = ws.Cells(i, 2).Value ' Date
            startOpen = ws.Cells(i, 3).Value ' Open price
            totalVolume = ws.Cells(i, 7).Value ' Volume

            ' Loop to find the end of the quarter for the same ticker
            For j = i + 1 To lastRow ' Start nested loop
                If ws.Cells(j, 1).Value <> ticker Then Exit For ' Exit loop if ticker changed
                If DateDiff("q", startDate, ws.Cells(j, 2).Value) > 0 Then Exit For ' Exit loop if quarter date changed

                endDate = ws.Cells(j, 2).Value ' Assigned end date
                endClose = ws.Cells(j, 6).Value ' Close price
                totalVolume = totalVolume + ws.Cells(j, 7).Value ' Accumulate volume for the same ticker
            Next j

            ' Calculate the quarterly and percent changes
            quarterlyChange = endClose - startOpen ' Calculate price change over the quarter
            If startOpen <> 0 Then
                percentChange = quarterlyChange / startOpen ' Calculate percent change
            Else
                percentChange = 0
            End If

            ' Aggregate the results for the ticker
            If Not dict.Exists(ticker) Then ' Check if ticker already exists in dictionary
                dict.Add ticker, CreateObject("Scripting.Dictionary")
                dict(ticker).Add "QuarterlyChange", 0 ' Initialize quarterly change
                dict(ticker).Add "PercentChange", 0 ' Initialize percent change
                dict(ticker).Add "TotalVolume", 0  ' Initialize total volume
            End If

            dict(ticker)("QuarterlyChange") = dict(ticker)("QuarterlyChange") + quarterlyChange ' Update dictionary
            dict(ticker)("PercentChange") = dict(ticker)("PercentChange") + percentChange
            dict(ticker)("TotalVolume") = dict(ticker)("TotalVolume") + totalVolume

            ' Move to the next row after processing the current quarter
            i = j - 1
        Next i

        ' Output the aggregated results
        ws.Cells(1, 8).Value = "Ticker"
        ws.Cells(1, 9).Value = "Quarterly Change"
        ws.Cells(1, 10).Value = "Percent Change"
        ws.Cells(1, 11).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"

        Counter = output
        
        For Each key In dict.Keys
            ws.Cells(Counter, 8).Value = key ' Aggregated Ticker
            ws.Cells(Counter, 9).Value = dict(key)("QuarterlyChange") ' Aggregated Quarterly Change
            ws.Cells(Counter, 10).Value = dict(key)("PercentChange") ' Aggregated Percent Change
            ws.Cells(Counter, 11).Value = dict(key)("TotalVolume") ' Aggregated Total Stock Volume
            ws.Cells(Counter, 10).NumberFormat = "0.00%"
            ws.Cells(2, 16).NumberFormat = "0.00%" ' Format as percentage
            ws.Cells(3, 16).NumberFormat = "0.00%"
        
            ' Max and min percent changes
            If dict(key)("PercentChange") > percentMax Then
                percentMax = dict(key)("PercentChange")
                tickerMaxPercent = key
            End If

            If dict(key)("PercentChange") < percentMin Then
                percentMin = dict(key)("PercentChange")
                tickerMinPercent = key
            End If

            ' Max total volume
            If dict(key)("TotalVolume") > maxVolume Then
                maxVolume = dict(key)("TotalVolume")
                tickerMaxVolume = key
            End If

            Counter = Counter + 1
            
        Next key

        ' Output to designated cells
        ws.Cells(2, 15).Value = tickerMaxPercent
        ws.Cells(2, 16).Value = percentMax
        ws.Cells(3, 15).Value = tickerMinPercent
        ws.Cells(3, 16).Value = percentMin
        ws.Cells(4, 15).Value = tickerMaxVolume
        ws.Cells(4, 16).Value = maxVolume

        ' Apply conditional formatting
        With ws.Range("I2:I" & Counter - 1).FormatConditions
            .Delete
            .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
            With .Item(.Count)
                .Interior.ColorIndex = 4 ' Green for positive numbers
            End With

            .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            With .Item(.Count)
                .Interior.ColorIndex = 3 ' Red for negative numbers
            End With
        End With

    Next ws

    MsgBox "Calculations completed!"

End Sub

