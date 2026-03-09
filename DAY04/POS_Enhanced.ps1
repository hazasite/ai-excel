# ============================================================
#   KOTTU KADE - ENHANCED POS SYSTEM  |  DAY04 v3
#   NEW: SET DISCOUNT popup + ENTER CASH -> Change popup
# ============================================================

$fileName = "KottuKade_POS_Enhanced.xlsm"
$filePath = Join-Path -Path $PSScriptRoot -ChildPath $fileName

if (Test-Path $filePath) { Remove-Item $filePath -Force -ErrorAction SilentlyContinue }

Clear-Host
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   KOTTU KADE POS v3  |  Generating..." -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Add()
$ws = $workbook.Worksheets.Item(1)
$ws.Name = "POS"

# ============================================================
# LAYOUT SETTINGS
# ============================================================
$excel.ActiveWindow.DisplayGridlines = $false
$excel.ActiveWindow.Zoom = 85

# Row heights - uniform
for ($r = 1; $r -le 50; $r++) { $ws.Rows.Item($r).RowHeight = 16 }

# Background - dark
$ws.Cells.Interior.Color = 0x1E1E2E

# Column widths
# A = pad | B C D E = buttons (4 cols) | F = divider | G H I J = bill | K = pad
$ws.Columns.Item(1).ColumnWidth = 1.5    # A
$ws.Columns.Item(2).ColumnWidth = 13.5   # B
$ws.Columns.Item(3).ColumnWidth = 13.5   # C
$ws.Columns.Item(4).ColumnWidth = 13.5   # D
$ws.Columns.Item(5).ColumnWidth = 13.5   # E
$ws.Columns.Item(6).ColumnWidth = 1.8    # F divider strip
$ws.Columns.Item(7).ColumnWidth = 24     # G description
$ws.Columns.Item(8).ColumnWidth = 6      # H qty
$ws.Columns.Item(9).ColumnWidth = 12     # I unit price
$ws.Columns.Item(10).ColumnWidth = 13     # J total
$ws.Columns.Item(11).ColumnWidth = 1.5    # K

# Orange divider strip
$ws.Range("F1:F50").Interior.Color = 0xF97316

# ============================================================
# LEFT PANEL - TITLE
# ============================================================
$ws.Rows.Item(1).RowHeight = 22
$ws.Range("B1:E1").Merge()
$ws.Cells.Item(1, 2).Value2 = "KOTTU KADE  POS"
$ws.Cells.Item(1, 2).Font.Bold = $true
$ws.Cells.Item(1, 2).Font.Size = 13
$ws.Cells.Item(1, 2).Font.Color = 0xF97316
$ws.Cells.Item(1, 2).Interior.Color = 0x12121E
$ws.Cells.Item(1, 2).HorizontalAlignment = -4108  # center

# Category label helper
function Set-CatLabel($row, $txt) {
    $ws.Rows.Item($row).RowHeight = 14
    $ws.Range("B${row}:E${row}").Merge()
    $ws.Cells.Item($row, 2).Value2 = $txt
    $ws.Cells.Item($row, 2).Font.Bold = $true
    $ws.Cells.Item($row, 2).Font.Size = 8
    $ws.Cells.Item($row, 2).Font.Color = 0xAAAAAA
    $ws.Cells.Item($row, 2).Interior.Color = 0x2A2A3E
    $ws.Cells.Item($row, 2).HorizontalAlignment = -4131  # left
}

# Category labels at specific rows
Set-CatLabel 2  "  KOTTU"
Set-CatLabel 6  "  RICE DISHES"
Set-CatLabel 10 "  NOODLES"
Set-CatLabel 14 "  DEVILLED"
Set-CatLabel 18 "  BEVERAGES"
Set-CatLabel 22 "  EXTRAS"
Set-CatLabel 26 "  ACTIONS"

# ============================================================
# BILL PANEL - WHITE INVOICE AREA
# COMPACT: rows 1-39, bill items rows 9-28 (20 rows max)
# ============================================================
$ws.Range("G1:J39").Interior.Color = 0xFFFFFF

# --- Header ---
$ws.Rows.Item(1).RowHeight = 22
$ws.Range("G1:J1").Merge()
$ws.Cells.Item(1, 7).Value2 = "KOTTU KADE RESTAURANT"
$ws.Cells.Item(1, 7).Font.Bold = $true
$ws.Cells.Item(1, 7).Font.Size = 13
$ws.Cells.Item(1, 7).Font.Color = 0x1E1E2E
$ws.Cells.Item(1, 7).HorizontalAlignment = -4108

$ws.Rows.Item(2).RowHeight = 13
$ws.Range("G2:J2").Merge()
$ws.Cells.Item(2, 7).Value2 = "No. 42, Galle Road, Colombo 03  |  Tel: 011-234-5678"
$ws.Cells.Item(2, 7).Font.Size = 7
$ws.Cells.Item(2, 7).Font.Color = 0x666666
$ws.Cells.Item(2, 7).HorizontalAlignment = -4108

# Orange line separator
$ws.Rows.Item(3).RowHeight = 3
$ws.Range("G3:J3").Interior.Color = 0xF97316

# --- Meta row (Invoice No / Date / Time / Table) ---
$ws.Rows.Item(4).RowHeight = 12
$ws.Rows.Item(5).RowHeight = 13
$ws.Cells.Item(4, 7).Value2 = "Invoice No:"
$ws.Cells.Item(4, 8).Value2 = "Date:"
$ws.Cells.Item(4, 9).Value2 = "Time:"
$ws.Cells.Item(4, 10).Value2 = "Table:"
for ($c = 7; $c -le 10; $c++) {
    $ws.Cells.Item(4, $c).Font.Bold = $true
    $ws.Cells.Item(4, $c).Font.Size = 7
    $ws.Cells.Item(4, $c).Font.Color = 0x444444
}

$ws.Cells.Item(5, 7).Formula = "=GetBillNo()"
$ws.Cells.Item(5, 8).Formula = '=TEXT(TODAY(),"DD/MM/YYYY")'
$ws.Cells.Item(5, 9).Formula = '=TEXT(NOW(),"HH:MM AM/PM")'
$ws.Cells.Item(5, 10).Value2 = "T-01"
for ($c = 7; $c -le 10; $c++) {
    $ws.Cells.Item(5, $c).Font.Size = 7
    $ws.Cells.Item(5, $c).Font.Color = 0x555555
}

# Grey separator
$ws.Rows.Item(6).RowHeight = 2
$ws.Range("G6:J6").Interior.Color = 0xCCCCCC

# --- Column headers ---
$ws.Rows.Item(7).RowHeight = 16
$ws.Cells.Item(7, 7).Value2 = "DESCRIPTION"
$ws.Cells.Item(7, 8).Value2 = "QTY"
$ws.Cells.Item(7, 9).Value2 = "UNIT PRICE"
$ws.Cells.Item(7, 10).Value2 = "TOTAL"
$ws.Range("G7:J7").Font.Bold = $true
$ws.Range("G7:J7").Font.Size = 8
$ws.Range("G7:J7").Font.Color = 0x1E1E2E
$ws.Range("G7:J7").Interior.Color = 0xF0F0F0
$ws.Range("G7:J7").HorizontalAlignment = -4108
$ws.Range("G7:J7").Borders.Item(9).LineStyle = 1
$ws.Range("G7:J7").Borders.Item(9).Color = 0xF97316
$ws.Range("G7:J7").Borders.Item(9).Weight = 3

# --- Item rows: 8 to 27 (20 rows) ---
$ws.Range("G8:J27").Interior.Color = 0xFFFFFF
for ($r = 8; $r -le 27; $r++) { $ws.Rows.Item($r).RowHeight = 15 }

# --- Subtraction area separator ---
$ws.Rows.Item(28).RowHeight = 2
$ws.Range("G28:J28").Interior.Color = 0xCCCCCC

# --- TOTALS: rows 29-36 ---
for ($r = 29; $r -le 36; $r++) { $ws.Rows.Item($r).RowHeight = 15 }

# Subtotal
$ws.Cells.Item(29, 9).Value2 = "SUBTOTAL:"
$ws.Cells.Item(29, 10).Formula = "=SUM(J8:J27)"

# Discount input
$ws.Range("G30:H30").Merge()
$ws.Cells.Item(30, 7).Value2 = "Discount %:"
$ws.Cells.Item(30, 7).Font.Size = 8
$ws.Cells.Item(30, 7).Font.Color = 0x555555
# Yellow input cell for discount
$ws.Cells.Item(30, 8).Value2 = 0
$ws.Cells.Item(30, 8).Interior.Color = 0xFFF9C4
$ws.Cells.Item(30, 8).Borders.Item(9).LineStyle = 1
$ws.Cells.Item(30, 8).Borders.Item(9).Color = 0xF97316
$ws.Cells.Item(30, 9).Value2 = "DISCOUNT:"
$ws.Cells.Item(30, 10).Formula = "=J29*(H30/100)"

# Service charge
$ws.Cells.Item(31, 9).Value2 = "SERVICE (10%):"
$ws.Cells.Item(31, 10).Formula = "=J29*0.10"

# Net Total - dark bg highlight
$ws.Rows.Item(32).RowHeight = 17
$ws.Range("G32:H32").Interior.Color = 0x1E1E2E
$ws.Range("I32:J32").Interior.Color = 0x1E1E2E
$ws.Cells.Item(32, 9).Value2 = "NET TOTAL (RS):"
$ws.Cells.Item(32, 10).Formula = "=J29-J30+J31"
$ws.Range("I32:J32").Font.Bold = $true
$ws.Range("I32:J32").Font.Size = 10
$ws.Range("I32:J32").Font.Color = 0xF97316

# Cash given - green input
$ws.Cells.Item(33, 9).Value2 = "CASH GIVEN:"
$ws.Cells.Item(33, 10).Value2 = 0
$ws.Cells.Item(33, 10).Interior.Color = 0xDFF0D8

# Change
$ws.Cells.Item(34, 9).Value2 = "CHANGE:"
$ws.Cells.Item(34, 10).Formula = "=J33-J32"
$ws.Cells.Item(34, 10).Font.Bold = $true
$ws.Cells.Item(34, 10).Font.Color = 0x155724

# Format all total rows
for ($r = 29; $r -le 34; $r++) {
    $ws.Cells.Item($r, 9).Font.Bold = $true
    $ws.Cells.Item($r, 9).Font.Size = 8
    $ws.Cells.Item($r, 9).Font.Color = 0x333333
    $ws.Cells.Item($r, 9).HorizontalAlignment = -4152  # xlRight
    $ws.Cells.Item($r, 10).NumberFormat = '"Rs." #,##0.00'
}

# Orange bottom line
$ws.Rows.Item(35).RowHeight = 3
$ws.Range("G35:J35").Interior.Color = 0xF97316

# Footer
$ws.Rows.Item(36).RowHeight = 13
$ws.Range("G36:J36").Merge()
$ws.Cells.Item(36, 7).Value2 = "Thank you for visiting!  wifi: kottu@free"
$ws.Cells.Item(36, 7).Font.Size = 7
$ws.Cells.Item(36, 7).Font.Color = 0x888888
$ws.Cells.Item(36, 7).HorizontalAlignment = -4108

# ============================================================
# VBA MACROS
# ============================================================
$vbaModule = $workbook.VBProject.VBComponents.Add(1)

$vbaCode = @"
Dim BillCounter As Long

Function GetBillNo() As String
    If BillCounter = 0 Then BillCounter = 1000
    GetBillNo = "KK-" & Format(BillCounter, "0000")
End Function

Sub AddItem(itemName As String, itemPrice As Double)
    Dim r As Long
    r = 8
    Do While r <= 27
        If Cells(r, 7).Value = itemName Then
            Cells(r, 8).Value = Cells(r, 8).Value + 1
            Exit Sub
        End If
        r = r + 1
    Loop

    r = 8
    Do While Cells(r, 7).Value <> "" And r <= 27
        r = r + 1
    Loop

    If r <= 27 Then
        Cells(r, 7).Value = itemName
        Cells(r, 8).Value = 1
        Cells(r, 9).Value = itemPrice
        Cells(r, 9).NumberFormat = """Rs."" #,##0.00"
        Cells(r, 10).Formula = "=H" & r & "*I" & r
        Cells(r, 10).NumberFormat = """Rs."" #,##0.00"
        With Range(Cells(r, 7), Cells(r, 10)).Borders(9)
            .LineStyle = 1
            .Weight = 2
            .Color = &HDDDDDD
        End With
        If (r Mod 2 = 0) Then
            Range(Cells(r, 7), Cells(r, 10)).Interior.Color = &HF7F7F7
        End If
    Else
        MsgBox "Bill is full! Please print and start a new bill.", vbCritical, "Bill Full"
    End If
End Sub

' -- KOTTU --
Sub AddChickenKottu()   : AddItem "Chicken Kottu", 1250    : End Sub
Sub AddEggKottu()       : AddItem "Egg Kottu", 950         : End Sub
Sub AddVegKottu()       : AddItem "Vegetable Kottu", 850   : End Sub
Sub AddFishKottu()      : AddItem "Fish Kottu", 1150       : End Sub

' -- RICE --
Sub AddFriedRice()      : AddItem "Chicken Fried Rice", 1100 : End Sub
Sub AddEggRice()        : AddItem "Egg Fried Rice", 900    : End Sub
Sub AddNasiGoreng()     : AddItem "Nasi Goreng", 1350      : End Sub
Sub AddBiriyani()       : AddItem "Chicken Biriyani", 1450 : End Sub

' -- NOODLES --
Sub AddChowMein()       : AddItem "Chow Mein", 1100        : End Sub
Sub AddFriedNoodles()   : AddItem "Fried Noodles", 950     : End Sub
Sub AddPadThai()        : AddItem "Pad Thai", 1250         : End Sub
Sub AddRicePaste()      : AddItem "Rice Paste Noodles", 900 : End Sub

' -- DEVILLED --
Sub AddDevChicken()     : AddItem "Devilled Chicken", 1500 : End Sub
Sub AddDevPork()        : AddItem "Devilled Pork", 1400    : End Sub
Sub AddDevSquid()       : AddItem "Devilled Squid", 1600   : End Sub
Sub AddDevVeg()         : AddItem "Devilled Veggies", 950  : End Sub

' -- BEVERAGES --
Sub AddCoke()           : AddItem "Coca-Cola (330ml)", 250 : End Sub
Sub AddWater()          : AddItem "Mineral Water", 100     : End Sub
Sub AddJuice()          : AddItem "Fresh Juice", 350       : End Sub
Sub AddTea()            : AddItem "Ceylon Tea", 150        : End Sub

' -- EXTRAS --
Sub AddPapad()          : AddItem "Papad", 100             : End Sub
Sub AddSambol()         : AddItem "Pol Sambol", 120        : End Sub
Sub AddFries()          : AddItem "French Fries", 500      : End Sub
Sub AddGarlicBread()    : AddItem "Garlic Bread", 350      : End Sub

' ============================================================
'  SET DISCOUNT - popup InputBox
' ============================================================
Sub SetDiscount()
    Dim subtotal As Double
    subtotal = Range("J29").Value
    If subtotal <= 0 Then
        MsgBox "Add items to the bill first!", vbExclamation, "No Items"
        Exit Sub
    End If

    Dim inp As String
    inp = InputBox( _
        "Current Subtotal:  Rs. " & Format(subtotal, "#,##0.00") & Chr(10) & Chr(10) & _
        "Enter Discount Percentage (e.g. 10 for 10%):", _
        "Set Discount", Range("H30").Value)

    If inp = "" Then Exit Sub

    Dim pct As Double
    pct = Val(inp)
    If pct < 0 Or pct > 100 Then
        MsgBox "Please enter a value between 0 and 100.", vbExclamation, "Invalid"
        Exit Sub
    End If

    Range("H30").Value = pct
    ActiveSheet.Calculate

    Dim discAmt As Double
    discAmt = subtotal * (pct / 100)
    MsgBox "Discount Applied!" & Chr(10) & Chr(10) & _
           "  Discount " & pct & "% :    Rs. " & Format(discAmt, "#,##0.00") & Chr(10) & _
           "  Net Total :    Rs. " & Format(Range("J32").Value, "#,##0.00"), _
           vbInformation, "Discount OK"
End Sub

' ============================================================
'  ENTER CASH - popup, then shows change
' ============================================================
Sub EnterCash()
    Dim netTotal As Double
    netTotal = Range("J32").Value
    If netTotal <= 0 Then
        MsgBox "No items in the bill yet!", vbExclamation, "Empty Bill"
        Exit Sub
    End If

    Dim inp As String
    inp = InputBox( _
        "Bill Total (Net):  Rs. " & Format(netTotal, "#,##0.00") & Chr(10) & Chr(10) & _
        "Enter cash given by customer (Rs.):", _
        "Cash Received", "")

    If inp = "" Then Exit Sub

    Dim cashGiven As Double
    cashGiven = Val(inp)

    If cashGiven < netTotal Then
        MsgBox "Cash given (Rs. " & Format(cashGiven, "#,##0.00") & ") is LESS than the total!" & Chr(10) & _
               "Short by:  Rs. " & Format(netTotal - cashGiven, "#,##0.00"), _
               vbCritical, "Insufficient Cash"
        Exit Sub
    End If

    ' Write cash to bill
    Range("J33").Value = cashGiven
    ActiveSheet.Calculate

    Dim change As Double
    change = cashGiven - netTotal

    ' Show change popup
    MsgBox "PAYMENT RECEIVED" & Chr(10) & _
           String(30, "-") & Chr(10) & _
           "  Net Total  :  Rs. " & Format(netTotal, "#,##0.00") & Chr(10) & _
           "  Cash Given :  Rs. " & Format(cashGiven, "#,##0.00") & Chr(10) & _
           String(30, "-") & Chr(10) & _
           "  CHANGE     :  Rs. " & Format(change, "#,##0.00"), _
           vbInformation, "Change Due"
End Sub

' ============================================================
'  PRINT BILL -> PDF on Desktop
' ============================================================
Sub PrintBill()
    If Range("J32").Value <= 0 Then
        MsgBox "No items in the bill!", vbExclamation, "Empty Bill"
        Exit Sub
    End If
    Dim desk As String
    desk = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    Dim fPath As String
    fPath = desk & "\KottuKade_" & Format(Now, "YYYYMMDD_HHMMSS") & ".pdf"
    ActiveSheet.PageSetup.PrintArea = "G1:J36"
    ActiveSheet.PageSetup.PaperSize = 9
    ActiveSheet.PageSetup.Zoom = False
    ActiveSheet.PageSetup.FitToPagesWide = 1
    ActiveSheet.PageSetup.FitToPagesTall = 1
    Range("G1:J36").ExportAsFixedFormat _
        Type:=0, Filename:=fPath, _
        Quality:=0, IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
    MsgBox "Invoice saved to Desktop!" & Chr(10) & fPath, vbInformation, "Printed OK"
End Sub

' ============================================================
'  NEW BILL (RESET)
' ============================================================
Sub NewBill()
    Dim ans As Integer
    ans = MsgBox("Clear this bill and start a new customer?", vbYesNo + vbQuestion, "New Bill")
    If ans <> vbYes Then Exit Sub

    With Range("G8:J27")
        .ClearContents
        .Interior.Color = &HFFFFFF
        .Borders.LineStyle = 0
    End With

    Range("H30").Value = 0
    Range("J33").Value = 0

    BillCounter = BillCounter + 1
    ActiveSheet.Calculate
    MsgBox "New bill ready!  " & GetBillNo(), vbInformation, "New Customer"
End Sub
"@

$vbaModule.CodeModule.AddFromString($vbaCode)

# ============================================================
# BUTTON HELPER
# ============================================================
$shapes = $ws.Shapes

function New-Btn($left, $top, $w, $h, $label, $macro, $bgRGB, $fgRGB) {
    $b = $shapes.AddShape(5, $left, $top, $w, $h)
    $b.TextFrame.Characters().Text = $label
    $b.TextFrame.Characters().Font.Bold = $true
    $b.TextFrame.Characters().Font.Size = 8
    $b.TextFrame.Characters().Font.Color = $fgRGB
    $b.TextFrame.HorizontalAlignment = -4108
    $b.TextFrame.VerticalAlignment = -4108
    $b.Fill.ForeColor.RGB = $bgRGB
    $b.OnAction = $macro
    $b.Line.Visible = 0
}

# Layout constants
# Each row of 4 buttons: x = 8, 78+8, 156+8, 234+8 => 8, 86, 164, 242
# Each button: w=72 h=24
$bw = 72; $bh = 24
$x1 = 8; $x2 = 82; $x3 = 156; $x4 = 230

# Row Y positions (matching category label rows)
# Cat labels are at rows 2,6,10,14,18,22 => pixel y approx
# Row 1 = 22px, row 2 = 14px, rows 3-5 = 16px each
# Accumulate: row1=0,row2=22, row3=36, row4=52,row5=68

# Kottu row: rows 3-5  (y start after row2 label)
$yKottu = 36
New-Btn $x1 $yKottu $bw $bh "Chicken Kottu"   "AddChickenKottu" 0xE05252 0xFFFFFF
New-Btn $x2 $yKottu $bw $bh "Egg Kottu"       "AddEggKottu"     0xE05252 0xFFFFFF
New-Btn $x3 $yKottu $bw $bh "Veg Kottu"       "AddVegKottu"     0xE05252 0xFFFFFF
New-Btn $x4 $yKottu $bw $bh "Fish Kottu"      "AddFishKottu"    0xE05252 0xFFFFFF
$yKottu2 = $yKottu + 28
New-Btn $x1 $yKottu2 $bw $bh "-- -- --"       ""                0x4A1A1A 0x4A1A1A  # spacer

# Rice row
$yRice = 36 + (4 * 16) + 14    # after kottu cat(14) + 4 rows (16ea): approx
# Simpler approach: calculate from row numbers
# row2=14, rows3-5=16ea=48, row6=14 => y6 start = 22+14+48 = 84
$yRice = 84
New-Btn $x1 $yRice $bw $bh "Egg Rice"         "AddEggRice"      0xE3A410 0x111111
New-Btn $x2 $yRice $bw $bh "Fried Rice"       "AddFriedRice"    0xE3A410 0x111111
New-Btn $x3 $yRice $bw $bh "Nasi Goreng"      "AddNasiGoreng"   0xE3A410 0x111111
New-Btn $x4 $yRice $bw $bh "Biriyani"         "AddBiriyani"     0xE3A410 0x111111

# Noodles row
# row6=14,rows7-9=48 => y10 = 22+14+48+14+48 = 146
$yNoodle = 146
New-Btn $x1 $yNoodle $bw $bh "Chow Mein"      "AddChowMein"     0x2D8A4E 0xFFFFFF
New-Btn $x2 $yNoodle $bw $bh "Fried Noodles"  "AddFriedNoodles" 0x2D8A4E 0xFFFFFF
New-Btn $x3 $yNoodle $bw $bh "Pad Thai"       "AddPadThai"      0x2D8A4E 0xFFFFFF
New-Btn $x4 $yNoodle $bw $bh "Rice Paste"     "AddRicePaste"    0x2D8A4E 0xFFFFFF

# Devilled row
$yDevil = 208
New-Btn $x1 $yDevil $bw $bh "Dev. Chicken"    "AddDevChicken"   0xCC2255 0xFFFFFF
New-Btn $x2 $yDevil $bw $bh "Dev. Pork"       "AddDevPork"      0xCC2255 0xFFFFFF
New-Btn $x3 $yDevil $bw $bh "Dev. Squid"      "AddDevSquid"     0xCC2255 0xFFFFFF
New-Btn $x4 $yDevil $bw $bh "Dev. Veg"        "AddDevVeg"       0xCC2255 0xFFFFFF

# Beverages row
$yBev = 270
New-Btn $x1 $yBev $bw $bh "Coke"              "AddCoke"         0x0080CC 0xFFFFFF
New-Btn $x2 $yBev $bw $bh "Water"             "AddWater"        0x0080CC 0xFFFFFF
New-Btn $x3 $yBev $bw $bh "Juice"             "AddJuice"        0x0080CC 0xFFFFFF
New-Btn $x4 $yBev $bw $bh "Tea"               "AddTea"          0x0080CC 0xFFFFFF

# Extras row
$yExtra = 332
New-Btn $x1 $yExtra $bw $bh "Papad"           "AddPapad"        0x7B3FA0 0xFFFFFF
New-Btn $x2 $yExtra $bw $bh "Pol Sambol"      "AddSambol"       0x7B3FA0 0xFFFFFF
New-Btn $x3 $yExtra $bw $bh "French Fries"    "AddFries"        0x7B3FA0 0xFFFFFF
New-Btn $x4 $yExtra $bw $bh "Garlic Bread"    "AddGarlicBread"  0x7B3FA0 0xFFFFFF

# ---- ACTION BUTTONS ----
# Row 1: SET DISCOUNT (amber) + ENTER CASH (green)
$yAction1 = 394
$bDisc = $shapes.AddShape(5, $x1, $yAction1, 148, 32)
$bDisc.TextFrame.Characters().Text = "SET DISCOUNT"
$bDisc.TextFrame.Characters().Font.Bold = $true
$bDisc.TextFrame.Characters().Font.Size = 10
$bDisc.TextFrame.Characters().Font.Color = 0x111111
$bDisc.TextFrame.HorizontalAlignment = -4108
$bDisc.TextFrame.VerticalAlignment = -4108
$bDisc.Fill.ForeColor.RGB = 0xF59E0B   # Amber
$bDisc.OnAction = "SetDiscount"
$bDisc.Line.Visible = 0

$bCash = $shapes.AddShape(5, ($x1 + 154), $yAction1, 148, 32)
$bCash.TextFrame.Characters().Text = "ENTER CASH"
$bCash.TextFrame.Characters().Font.Bold = $true
$bCash.TextFrame.Characters().Font.Size = 10
$bCash.TextFrame.Characters().Font.Color = 0xFFFFFF
$bCash.TextFrame.HorizontalAlignment = -4108
$bCash.TextFrame.VerticalAlignment = -4108
$bCash.Fill.ForeColor.RGB = 0x16A34A   # Dark green
$bCash.OnAction = "EnterCash"
$bCash.Line.Visible = 0

# Row 2: PRINT INVOICE (teal) + NEW CUSTOMER (blue)
$yAction2 = 432
$bPrint = $shapes.AddShape(5, $x1, $yAction2, 148, 32)
$bPrint.TextFrame.Characters().Text = "PRINT  INVOICE"
$bPrint.TextFrame.Characters().Font.Bold = $true
$bPrint.TextFrame.Characters().Font.Size = 10
$bPrint.TextFrame.Characters().Font.Color = 0xFFFFFF
$bPrint.TextFrame.HorizontalAlignment = -4108
$bPrint.TextFrame.VerticalAlignment = -4108
$bPrint.Fill.ForeColor.RGB = 0x0891B2   # Teal
$bPrint.OnAction = "PrintBill"
$bPrint.Line.Visible = 0

$bReset = $shapes.AddShape(5, ($x1 + 154), $yAction2, 148, 32)
$bReset.TextFrame.Characters().Text = "NEW  CUSTOMER"
$bReset.TextFrame.Characters().Font.Bold = $true
$bReset.TextFrame.Characters().Font.Size = 10
$bReset.TextFrame.Characters().Font.Color = 0xFFFFFF
$bReset.TextFrame.HorizontalAlignment = -4108
$bReset.TextFrame.VerticalAlignment = -4108
$bReset.Fill.ForeColor.RGB = 0x1D4ED8   # Blue
$bReset.OnAction = "NewBill"
$bReset.Line.Visible = 0

# ============================================================
# FREEZE PANES & SAVE
# ============================================================
# Freeze at row 8 col G so headers stay
$excel.ActiveWindow.FreezePanes = $false
$ws.Activate()

$workbook.SaveAs($filePath, 52)

Write-Host ""
Write-Host "  Done! -> $fileName" -ForegroundColor Green
Write-Host ""
Write-Host "  NEW in v3:" -ForegroundColor Yellow
Write-Host "   [x] SET DISCOUNT button -> popup asks % -> shows discount Rs." -ForegroundColor White
Write-Host "   [x] ENTER CASH button  -> popup asks cash -> shows CHANGE" -ForegroundColor White
Write-Host "   [x] Insufficient cash warning popup" -ForegroundColor White
Write-Host "   [x] Cash less than total -> error popup" -ForegroundColor White
Write-Host "   [x] 4 action buttons: Discount / Cash / Print / New" -ForegroundColor White
Write-Host ""
