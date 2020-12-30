{\rtf1\ansi\ansicpg1252\cocoartf2577
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub stockTicker_main()\
\
'declare variables\
Dim ticker As String\
Dim open_price, close_price, total_volume As Double\
Dim summary_table_row As Integer\
Dim ws As Worksheet\
\
'turn off screen updating to make the process faster\
Application.ScreenUpdating = False\
\
'loop through each of the worksheets\
For Each ws In Worksheets\
    Worksheets(ws.Name).Activate\
    'sort the spreadsheet by stock ticker and then by year\
    With ActiveSheet.Sort\
         .SortFields.Add Key:=Range("A2"), Order:=xlAscending\
         .SortFields.Add Key:=Range("B2"), Order:=xlAscending\
         .SetRange Range("A2", Range("G2").End(xlDown))\
         .Apply\
    End With\
    \
    'assign initial values to variables\
    'columns I, J, K and L are the header rows for the initial project\
    Range("I1").Value = "Ticker"\
    Range("J1").Value = "Yearly Change"\
    Range("K1").Value = "Percent Change"\
    Range("L1").Value = "Total Stock Volume"\
    open_price = Cells(2, 3).Value\
    total_volume = 0\
    summary_table_row = 2\
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row\
    \
    'loop through the stock data, by ticker, by year\
    'calculate the difference between open/close price by stock ticker\
    'calculate the change in price and % change\
    \
    For i = 2 To lastrow\
    \
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then 'this evaluates true when the tickers change in column A\
            ticker = Cells(i, 1).Value\
            total_volume = total_volume + Cells(i, 7).Value\
            close_price = Cells(i, 6).Value\
            Range("I" & summary_table_row).Value = ticker\
            Range("L" & summary_table_row).Value = total_volume\
            next_ticker_first_row = i + 1 'save the first row of the next ticker for the open price\
            yearly_change = close_price - open_price\
            Range("J" & summary_table_row).Value = Format(yearly_change, "Fixed")\
            If open_price = 0 Then 'adjust for $0.00 open price to avoid division by zero error\
                percent_change = 0\
            Else\
                percent_change = close_price / open_price - 1\
            End If\
            Range("K" & summary_table_row).Value = Format(percent_change, "Percent")\
                If yearly_change < 0 Then 'if the annual close price is less than the open price, fill the cell red\
                    Range("J" & summary_table_row).Interior.ColorIndex = 3\
                Else 'if the annual close price is higher than the open price, fill the cell green\
                    Range("J" & summary_table_row).Interior.ColorIndex = 4\
                End If\
            open_price = Cells(next_ticker_first_row, 3).Value 'set the open price at each change in ticker\
            total_volume = 0 'reset the the total volume at each change in ticker\
            summary_table_row = summary_table_row + 1\
        Else\
            total_volume = total_volume + Cells(i, 7).Value 'keep adding to the total until the ticker changes\
        End If\
    \
    Next i\
    \
    'best fit each of the columns so it's easier for the user to read\
    Range("A:Q").EntireColumn.AutoFit\
Next ws\
\
'turn screen updating back on\
Application.ScreenUpdating = True\
\
'display a message box so the user knows the calculations are finished\
MsgBox (\'93Main Homework Complete")\
\
End Sub\
\
}