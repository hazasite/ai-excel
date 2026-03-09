# 1. Path Setup
$fileName = "Kottu_Standard_Final_V3.xlsm"
$filePath = Join-Path -Path $PSScriptRoot -ChildPath $fileName

if (Test-Path $filePath) { Remove-Item $filePath -Force -ErrorAction SilentlyContinue }

Clear-Host
Write-Host "--- GENERATING OPTIMIZED KOTTU POS ---" -ForegroundColor Cyan

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Add()
$ws = $workbook.Worksheets.Item(1)
$ws.Name = "POS_Terminal"

# 2. UI Layout Setup
$excel.ActiveWindow.DisplayGridlines = $false
$ws.Cells.Interior.ColorIndex = 15 # Light Grey Background

# DESIGNING THE WHITE BILL AREA
$ws.Range("G2:K40").Interior.ColorIndex = 2 # Pure White
$ws.Columns.Item(7).ColumnWidth = 25 # Description
$ws.Columns.Item(8).ColumnWidth = 8  # Qty
$ws.Columns.Item(9).ColumnWidth = 12 # Price
$ws.Columns.Item(10).ColumnWidth = 15 # Total

# Standard Invoice Header
$ws.Range("G3:J3").Merge(); $ws.Cells.Item(3, 7) = "KOTTU KADE RESTAURANT"
$ws.Cells.Item(3, 7).Font.Size = 14; $ws.Cells.Item(3, 7).Font.Bold = $true; $ws.Cells.Item(3, 7).HorizontalAlignment = -4108

$ws.Range("G4:J4").Merge(); $ws.Cells.Item(4, 7) = "Standard Invoice - Colombo 03"
$ws.Cells.Item(4, 7).Font.Size = 9; $ws.Cells.Item(4, 7).HorizontalAlignment = -4108

# Headers for items
$ws.Cells.Item(9, 7) = "DESCRIPTION"; $ws.Cells.Item(9, 8) = "QTY"; $ws.Cells.Item(9, 9) = "PRICE"; $ws.Cells.Item(9, 10) = "TOTAL"
$ws.Range("G9:J9").Font.Bold = $true; $ws.Range("G9:J9").Borders.Item(9).LineStyle = 1

# Grand Total (Fixed at Row 42)
$ws.Cells.Item(42, 9) = "NET TOTAL (RS):"
$ws.Cells.Item(42, 10).Formula = "=SUM(J10:J41)"
$ws.Range("I42:J42").Font.Bold = $true

# ==========================================
# 3. VBA MACRO (DYNAMIC DATA ENTRY)
# ==========================================
$vbaModule = $workbook.VBProject.VBComponents.Add(1)
$vbaCode = @"
Sub AddItem(name As String, price As Double)
    Dim r As Long
    ' Find the first truly empty row between Row 10 and 41
    r = 10
    Do While Cells(r, 7).Value <> "" And r < 41
        r = r + 1
    Loop
    
    If r <= 41 Then
        Cells(r, 7).Value = name
        Cells(r, 8).Value = 1
        Cells(r, 9).Value = price
        Cells(r, 10).Formula = "=H" & r & "*I" & r
        ' Add subtle border to row
        Range(Cells(r, 7), Cells(r, 10)).Borders.Item(9).LineStyle = 1
        Range(Cells(r, 7), Cells(r, 10)).Borders.Item(9).Weight = 2
    Else
        MsgBox "Inovice Page Full!", vbCritical
    End If
End Sub

Sub AddChicken() : AddItem "Chicken Kottu", 1250 : End Sub
Sub AddEgg()     : AddItem "Egg Kottu", 950 : End Sub
Sub AddRice()    : AddItem "Fried Rice", 1100 : End Sub

Sub PrintBill()
    Dim fPath As String
    fPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\Kottu_Invoice.pdf"
    Range("G2:J43").ExportAsFixedFormat Type:=0, Filename:=fPath
    MsgBox "Standard Invoice Saved to Desktop!", vbInformation
End Sub

Sub NewBill()
    Range("G10:J41").ClearContents
    Range("G10:J41").Borders.LineStyle = 0
End Sub
"@
$vbaModule.CodeModule.AddFromString($vbaCode)

# ==========================================
# 4. STYLED BUTTONS (Left Panel)
# ==========================================
$shapes = $ws.Shapes

function Create-Btn($left, $top, $width, $height, $text, $macro, $color) {
    $b = $shapes.AddShape(5, $left, $top, $width, $height) # Rounded Rect
    $b.TextFrame.Characters().Text = $text
    $b.TextFrame.Characters().Font.Bold = $true
    $b.TextFrame.Characters().Font.ColorIndex = 2
    $b.Fill.ForeColor.RGB = $color
    $b.OnAction = $macro
    $b.Line.Visible = 0
}

# Add Item Buttons
Create-Btn 40 60 160 40 "CHICKEN KOTTU" "AddChicken" 0x2E2E2E # Dark Grey
Create-Btn 40 110 160 40 "EGG KOTTU" "AddEgg" 0x2E2E2E
Create-Btn 40 160 160 40 "FRIED RICE" "AddRice" 0x2E2E2E

# Action Buttons
Create-Btn 40 240 160 50 "PRINT INVOICE" "PrintBill" 0x008000 # Green
Create-Btn 40 300 160 35 "NEW CUSTOMER" "NewBill" 0x0000FF # Blue

# 5. Save and Close
$workbook.SaveAs($filePath, 52)
Write-Host "Success! Use the POS Panel on the left." -ForegroundColor Green