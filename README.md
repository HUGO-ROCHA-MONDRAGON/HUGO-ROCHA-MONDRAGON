 I’m @HUGO-ROCHA-MONDRAGON, student at Université Paris Dauphine-PSL (Financial Engineering).
Sub InsertCouponPayments()
    Dim ws As Worksheet
    Dim lastCoupon As Date
    Dim maturity As Date
    Dim frequency As Integer
    Dim nominal As Double
    Dim coupon As Double
    Dim period As Double
    Dim nextDate As Date
    Dim col As Range
    
    ' Set the worksheet (modify if needed)
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your actual sheet name

    ' User inputs
    lastCoupon = DateValue("06/12/2025") ' Last coupon date (dd/mm/yyyy)
    maturity = DateValue("06/12/2030") ' Maturity date
    frequency = 2 ' Coupons per year (Annual=1, Semi-Annual=2, Quarterly=4)
    nominal = 1000 ' Nominal bond value
    coupon = nominal * 0.05 / frequency ' Example: 5% coupon rate (adjust if needed)

    ' Determine the period in months based on frequency
    period = 12 / frequency

    ' Start from the first coupon date
    nextDate = lastCoupon

    ' Loop through all coupon dates until maturity
    Do While nextDate <= maturity
        ' Find the column where the date is located in Row 1
        Set col = ws.Rows(1).Find(What:=nextDate, LookAt:=xlWhole)
        
        If Not col Is Nothing Then
            ' Place the coupon value in Row 2 of the matching column
            ws.Cells(2, col.Column).Value = coupon
        End If
        
        ' Move to the next coupon date
        nextDate = DateAdd("m", period, nextDate)
    Loop
    
    MsgBox "Coupons inserted successfully!", vbInformation
End Sub
<!---
HUGO-ROCHA-MONDRAGON/HUGO-ROCHA-MONDRAGON is a ✨ special ✨ repository because its `README.md` (this file) appears on your GitHub profile.
You can click the Preview link to take a look at your changes.
--->
