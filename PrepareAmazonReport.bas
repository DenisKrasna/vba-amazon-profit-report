' Module: PrepareAmazonReport
' Author: Denis Krašna
' Description: Calculates profit for an Amazon product using input price, cost, and commission.

Public wb As Workbook
Public ws As Worksheet, wsNew As Worksheet
Public sheetName As String, label1 As String, label2 As String, label3 As String, label4 As String
Public exists As Boolean
Public sellingPrice As Double, productionCost As Double, commission As Double, profit As Double

Sub GenerateAmazonReport()

    Call InitializeSheet
    Call InputProductData
    Call CalculateProfit
    Call ShowSummary

    Debug.Print "GenerateAmazonReport() completed successfully."

End Sub

Sub InitializeSheet()

    Set wb = Workbooks("porocilo.xlsm")

    sheetName = "AmazonReport"
    exists = False

    For Each ws In wb.Worksheets
        If ws.Name = sheetName Then
            exists = True
            Exit For
        End If
    Next ws

    If exists Then
        wb.Worksheets(sheetName).Range("A1:G10").ClearContents
        Debug.Print "Existing sheet [" & ws.Name & "] cleared."
    Else
        Set wsNew = wb.Worksheets.Add(before:=wb.Worksheets(1))
        wsNew.Name = sheetName
        Debug.Print "New sheet [" & wsNew.Name & "] created."
    End If

    Debug.Print "InitializeSheet() completed."

End Sub

Sub InputProductData()

    Dim inputSellingPrice As Double, inputProductionCost As Double
    inputSellingPrice = InputBox("Enter the product selling price:")
    inputProductionCost = InputBox("Enter the production cost of the product:")

    sellingPrice = CDbl(inputSellingPrice)
    productionCost = CDbl(inputProductionCost)

    If Not IsNumeric(sellingPrice) Or Not IsNumeric(productionCost) Then
        MsgBox "Invalid input for selling price or production cost."
        Exit Sub
    End If

    label1 = "Selling Price"
    label2 = "Production Cost"

    With wsNew
        .Range("A1").Value = label1
        .Range("A2").Value = label2
        .Range("B1").Value = sellingPrice
        .Range("B2").Value = productionCost
    End With

    Debug.Print "InputProductData() completed."

End Sub

Sub CalculateProfit()

    Dim inputCommission As Double
    inputCommission = InputBox("Enter the Amazon commission:")

    commission = CDbl(inputCommission)

    If Not IsNumeric(commission) Then
        MsgBox "Commission must be a numeric value."
        Exit Sub
    End If

    profit = sellingPrice - productionCost - commission

    label3 = "Amazon Commission"
    label4 = "Profit"

    With wsNew
        .Range("A3").Value = label3
        .Range("A4").Value = label4
        .Range("B3").Value = commission
        .Range("B4").Value = profit
    End With

    Debug.Print "CalculateProfit() completed."

End Sub

Sub ShowSummary()

    MsgBox label1 & vbTab & FormatCurrency(sellingPrice, 2) & vbCrLf & _
           label2 & vbTab & FormatCurrency(productionCost, 2) & vbCrLf & _
           label3 & vbTab & FormatCurrency(commission, 2) & vbCrLf & _
           label4 & vbTab & FormatCurrency(profit, 2)

    Debug.Print "ShowSummary() completed."

End Sub
