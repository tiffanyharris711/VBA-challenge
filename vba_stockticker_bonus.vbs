{\rtf1\ansi\ansicpg1252\cocoartf2577
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub stockTicker_bonus()\
\
'declare variables\
Dim ws As Worksheet\
\
'turn off screen updating to make the process faster\
Application.ScreenUpdating = False\
\
'loop through each of the worksheets\
For Each ws In Worksheets\
    Worksheets(ws.Name).Activate\
\
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row\
    \
    '**______BONUS START_____**\
    \
    'columns O, P and Q are the labels/headers for the bonus project\
    Range("O2").Value = "Greatest % Increase"\
    Range("O3").Value = "Greastest % Decrease"\
    Range("O4").Value = "Greastest Total Volume"\
    Range("P1").Value = "Ticker"\
    Range("Q1").Value = "Value"\
    'set the initial values to 0 for the comparison of greatest change\
    Range("Q2:Q4").Value = 0\
    greatest_percent_increase = 0\
    greatest_percent_decrease = 0\
    greatest_total_volume = 0\
    \
    For j = 2 To lastrow\
        'compare the current row value of percent change to the last saved value for greatest increase\
        If Cells(j, 11).Value > greatest_percent_increase Then\
            greatest_percent_increase_ticker = Cells(j, 9).Value\
            greatest_percent_increase = Cells(j, 11).Value\
            Range("P2").Value = greatest_percent_increase_ticker\
            Range("Q2").Value = Format(greatest_percent_increase, "Percent")\
        End If\
        'compare the current row value of percent change to the last saved value for greatest decrease\
        If Cells(j, 11).Value < greatest_percent_decrease Then\
            greatest_percent_decrease_ticker = Cells(j, 9).Value\
            greatest_percent_decrease = Cells(j, 11).Value\
            Range("P3").Value = greatest_percent_decrease_ticker\
            Range("Q3").Value = Format(greatest_percent_decrease, "Percent")\
        End If\
        'compare the current row value of total volume to the last saved value\
        If Cells(j, 12).Value > greatest_total_volume Then\
            greatest_total_volume_ticker = Cells(j, 9).Value\
            greatest_total_volume = Cells(j, 12).Value\
            Range("P4").Value = greatest_total_volume_ticker\
            Range("Q4").Value = greatest_total_volume\
        End If\
    Next j\
    \
    '**______BONUS END_____**\
    \
    'best fit each of the columns so it's easier for the user to read\
    Range("A:Q").EntireColumn.AutoFit\
Next ws\
\
'turn screen updating back on\
Application.ScreenUpdating = True\
\
'display a message box so the user knows the calculations are finished\
MsgBox ("Bonus Complete")\
\
End Sub\
\
}