# VBA-Challenge
'Name the assignment
Sub Stock_Analysis()

'Set ExcelWs as a worksheet Variable

Dim ExcelWs As Worksheet

For Each ExcelWs In Worksheets

'Label Column and Rows for Tables

ExcelWs.Range("J1").Value = "Ticker"
ExcelWs.Range("K1").Value = "Yearly Change"
ExcelWs.Range("L1").Value = "Percent Change"
ExcelWs.Range("M1").Value = "Volume"

ExcelWs.Range("O2").Value = "Greatest%Increase"
ExcelWs.Range("O3").Value = "Greatest%Decrease"
ExcelWs.Range("O4").Value = "Greatest Total Volume"

ExcelWs.Range("P1").Value = "Ticker"
ExcelWs.Range("Q1").Value = "Value"

'Label the String and Integer Variable I'll be working with

Dim Ticker_Name As String
Ticker_Name = " "

Dim Yearly_Change As Double
Yearly_Change = 0

Dim Percent_Change As Double
Percent_Change = 0

Dim Total_Volume As Double
Total_Volume = 0

'Label Second Table Strings and Integers

Dim Best_Yearly_Ticker As String
Best_Yearly_Ticker = " "

Dim Worst_Yearly_Ticker As String
Worst_Yearly_Ticker = " "

Dim Largest_Total_Ticker As String
Largest_Total_Ticker = " "

Dim Best_Yearly_Percent As Double
Best_Yearly_Percent = 0

Dim Worst_Yearly_Percent As Double
Worst_Yearly_Percent = 0

Dim Largest_Total_Volume As Double
Largest_Total_Volume = 0

'Create Variable for Table so that it loops to next row

Dim Row As Double
Row = 2

'Create Variables to calculate info needed for both tabels
'Columns B, D, E are not needed for calculations

Dim Open_Price As Double
Open_Price = 0

Dim Close_Price As Double
Close_Price = 0

'Counts the last row
Lastrow = ExcelWs.Cells(Rows.Count, 1).End(xlUp).Row

'Create starting value for Open Price and begin looping through rows

Open_Price = ExcelWs.Cells(2, 3).Value

For i = 2 To Lastrow
    If ExcelWs.Cells(i + 1, 1).Value <> ExcelWs.Cells(i, 1).Value Then

        Close_Price = ExcelWs.Cells(i, 6).Value
'Calculate Yearly Change
        Yearly_Change = Close_Price - Open_Price
'Calculate Total Volume
        Total_Volume = Total_Volume + ExcelWs.Cells(i, 7).Value
'Get the Ticker Name
        Ticker_Name = ExcelWs.Cells(i, 1).Value

'Calculate Percent Change
        Percent_Change = (Yearly_Change / Open_Price) * 100

'Place information into first summary table

ExcelWs.Range("J" & Row).Value = Ticker_Name

'Include If statement after inserting Yearly_Change to change the cells color depending on it's value
ExcelWs.Range("K" & Row).Value = Yearly_Change

        If Yearly_Change > 0 Then
        ExcelWs.Range("K" & Row).Interior.ColorIndex = 4
            Else
                ExcelWs.Range("K" & Row).Interior.ColorIndex = 3
                
        End If
ExcelWs.Range("L" & Row).Value = (CStr(Percent_Change) + "%")

ExcelWs.Range("M" & Row).Value = Total_Volume

'We're done with our first row in the summary table so we now need to add a row so that the next ticker doesn't overwrite the previous ticker

Row = Row + 1

'Start adding info into second small summary table
        If Percent_Change > Best_Yearly_Percent Then
           Best_Yearly_Percent = Percent_Change
           Best_Yearly_Ticker = Ticker_Name
           
           ElseIf Percent_Change <= Worst_Yearly_Percent Then
                  Worst_Yearly_Percent = Percent_Change
                  Worst_Yearly_Ticker = Ticker_Name

        End If

        If Total_Volume > Largest_Total_Volume Then
            Largest_Total_Volume = Total_Volume
            Largest_Total_Ticker = Ticker_Name

            End If
'Now I'll reset certain variables to calculate new values for the next ticker
Open_Price = ExcelWs.Cells(i + 1, 3).Value
Close_Price = 0
Yearly_Change = 0
Percent_Change = 0
Total_Volume = 0

'If the next row in our Excel Sheet happens to be equal to the next row then the code above won't run so I should have an else statement to continue adding the Volumes,
'Until my first If statement happens to be true
Else
    Total_Volume = Total_Volume + ExcelWs.Cells(i, 7).Value



'Close the very first If statement I made and move to the next i

End If

Next i

'After the code loops through the whole sheet, we can take the Best, Worst, and Largest Total and now place it in the Second table we created

ExcelWs.Range("P2").Value = Best_Yearly_Ticker
ExcelWs.Range("Q2").Value = (CStr(Best_Yearly_Percent) + "%")

ExcelWs.Range("P3").Value = Worst_Yearly_Ticker
ExcelWs.Range("Q3").Value = (CStr(Worst_Yearly_Percent) + "%")

ExcelWs.Range("P4").Value = Largest_Total_Ticker
ExcelWs.Range("Q4").Value = Largest_Total_Volume

Next ExcelWs

End Sub
