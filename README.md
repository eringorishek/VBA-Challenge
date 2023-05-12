# VBA-Challenge
# The attached screenshots and files show the work completed for Module 2 - VBA scripting. It includes screenshots of my results as well as the VBA scripting used to achieve the results. 
```vba
    
    Sub challenge2()

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        Dim last_row As Long
        Dim i As Long
        Dim j As Long
        Dim current_ticker As String
        Dim total_volume As Variant
        Dim year_open As Currency
        Dim year_close As Currency
        Dim greatest_percent_increase As Long
        Dim greatest_percent_decrease As Long
        Dim largest_volume As Long

        last_row = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        current_ticker = ws.Cells(2, "A").Value

        year_open = ws.Cells(2, "C").Value

        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"

        total_volume = 0
        j = 2

        For i = 2 To last_row
            If current_ticker <> ws.Cells(i, "A").Value Then
                'print out final stats for current ticker
                ws.Cells(j, "I").Value = current_ticker
                ws.Cells(j, "L").Value = total_volume
                total_volume = 0
                ws.Cells(j, "J").Value = year_close - year_open
                ws.Cells(j, "K").Value = (year_close - year_open) / year_open
                If ws.Cells(j, "J").Value < 0 Then
                    ws.Cells(j, "J").Interior.ColorIndex = 3
                Else: ws.Cells(j, "J").Interior.ColorIndex = 4
                End If
                  If ws.Cells(j, "K").Value < 0 Then
            ws.Cells(j, "K").Interior.ColorIndex = 3
        Else: ws.Cells(j, "K").Interior.ColorIndex = 4
        End If
                j = j + 1
            Else: year_close = ws.Cells(i, "D").Value
            End If

            current_ticker = ws.Cells(i, "A").Value
            total_volume = total_volume + ws.Cells(i, "G").Value

        Next i
        'print out final stats for last ticker
        ws.Cells(j, "I").Value = current_ticker
        ws.Cells(j, "L").Value = total_volume
        ws.Cells(j, "J").Value = year_close - year_open
        ws.Cells(j, "K").Value = (year_close - year_open) / year_open

        If ws.Cells(j, "J").Value < 0 Then
            ws.Cells(j, "J").Interior.ColorIndex = 3
        Else: ws.Cells(j, "J").Interior.ColorIndex = 4
        End If
          If ws.Cells(j, "K").Value < 0 Then
            ws.Cells(j, "K").Interior.ColorIndex = 3
        Else: ws.Cells(j, "K").Interior.ColorIndex = 4
        End If


        'part 2
        greatest_percent_increase = 2
        greatest_percent_decrease = 2
        largest_volume = 2

        For i = 3 To j
            If ws.Cells(i, "J").Value > ws.Cells(greatest_percent_increase, "J").Value Then
                greatest_percent_increase = i
            End If

            If ws.Cells(i, "J").Value < ws.Cells(greatest_percent_decrease, "J").Value Then
                greatest_percent_decrease = i
            End If

            If ws.Cells(i, "J").Value > ws.Cells(largest_volume, "J").Value Then
                largest_volume = i
            End If

        Next i

        ws.Cells(1, "O").Value = "Ticker"
        ws.Cells(1, "P").Value = "Value"
        ws.Cells(2, "N").Value = "Greatest % Increase"
        ws.Cells(3, "N").Value = "Greatest % Decrease"
        ws.Cells(4, "N").Value = "Greatest Total Volume"
        ws.Cells(2, "O").Value = ws.Cells(greatest_percent_increase, "I")
        ws.Cells(2, "P").Value = ws.Cells(greatest_percent_increase, "K")
        ws.Cells(3, "O").Value = ws.Cells(greatest_percent_decrease, "I")
        ws.Cells(3, "P").Value = ws.Cells(greatest_percent_decrease, "K")
        ws.Cells(4, "O").Value = ws.Cells(largest_volume, "I")
        ws.Cells(4, "P").Value = ws.Cells(largest_volume, "L")

    Next ws



    End Sub




```

<img width="1440" alt="Screenshot 2023-05-11 at 9 46 17 PM" src="https://github.com/eringorishek/VBA-Challenge/assets/130519405/9728304c-30a7-44e0-a582-27462d563c81">
<img width="1440" alt="Screenshot 2023-05-11 at 9 57 04 PM" src="https://github.com/eringorishek/VBA-Challenge/assets/130519405/e212e255-5f1e-4c67-8649-d96c4bcb3e1f">
<img width="1440" alt="Screenshot 2023-05-11 at 9 57 15 PM" src="https://github.com/eringorishek/VBA-Challenge/assets/130519405/4b7dc6a2-1795-485f-88b0-372aaccb1bc4">
<img width="1440" alt="Screenshot 2023-05-11 at 9 46 26 PM" src="https://github.com/eringorishek/VBA-Challenge/assets/130519405/74dc2eca-7b34-493a-8d8f-7e3b95b031b1">


