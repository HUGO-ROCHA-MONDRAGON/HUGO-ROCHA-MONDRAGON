 Sub Build_AB_C_From_Filters()

    Dim wsAAA As Worksheet, wsOUT As Worksheet
    Dim wsAuto As Worksheet, wsRMBS As Worksheet
    Dim lastrow As Long, lastOUT As Long
    Dim dict As Object
    Dim rng As Range, rFilt As Range
    
    ' Dictionnaire pour stocker Ticker + Description
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' === Feuilles sources ===
    Set wsAAA = ThisWorkbook.Worksheets("AAA_SSA")
    Set wsAuto = ThisWorkbook.Worksheets("Auto")
    Set wsRMBS = ThisWorkbook.Worksheets("RMBS")
    
    ' === Création feuille résultat ===
    Set wsOUT = ThisWorkbook.Worksheets.Add
    wsOUT.Name = "RESULT"
    
    ' Titres
    wsOUT.Range("A1").Value = "Ticker"
    wsOUT.Range("B1").Value = "Description"
    wsOUT.Range("C1").Value = "Imported Tickers (AUTO + RMBS)"
    
    ' --------------------------------------------------------------------
    '   1) FILTRER AAA_SSA → Auto Loans + Auto Leases + Prime RMBS + NC RMBS
    ' --------------------------------------------------------------------

    If wsAAA.AutoFilterMode Then wsAAA.AutoFilterMode = False
    
    lastrow = wsAAA.Cells(wsAAA.Rows.Count, 1).End(xlUp).Row
    
    ' 4 critères sur une seule colonne = on passe par un Array
    wsAAA.Range("A1").AutoFilter Field:=6, _
        Criteria1:=Array("Auto Loans", "Auto Leases", "Prime RMBS", "NC RMBS"), _
        Operator:=xlFilterValues
    
    ' Récupère la colonne B (Ticker)
    On Error Resume Next
    Set rFilt = wsAAA.Range("B2:B" & lastrow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not rFilt Is Nothing Then
        rFilt.Copy wsOUT.Range("A2")
    End If
    
    ' Récupère la colonne F (Description)
    On Error Resume Next
    Set rFilt = wsAAA.Range("F2:F" & lastrow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not rFilt Is Nothing Then
        rFilt.Copy wsOUT.Range("B2")
    End If
    
    ' Enlever le filtre
    If wsAAA.AutoFilterMode Then wsAAA.AutoFilterMode = False
    
    ' --------------------------------------------------------------------
    '   2) CONSTRUIRE LA COLONNE C AVEC FEUILLES AUTO & RMBS
    ' --------------------------------------------------------------------
    
    Dim arrSheets As Variant
    Dim sh As Variant
    
    arrSheets = Array(wsAuto, wsRMBS)
    
    For Each sh In arrSheets
        
        ' retire tout filtre
        If sh.AutoFilterMode Then sh.AutoFilterMode = False
        
        lastrow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
        
        ' met chaque ticker de la colonne A dans le dictionnaire
        For i = 2 To lastrow
            If sh.Cells(i, 1).Value <> "" Then
                dict(sh.Cells(i, 1).Value) = 1
            End If
        Next i
        
    Next sh
    
    ' --------------------------------------------------------------------
    '   3) ÉCRIRE LE DICTIONNAIRE → COLONNE C DE RESULT
    ' --------------------------------------------------------------------
    
    lastOUT = wsOUT.Cells(wsOUT.Rows.Count, 1).End(xlUp).Row + 1
    
    Dim key As Variant
    Dim writeRow As Long: writeRow = 2
    
    For Each key In dict.Keys
        wsOUT.Cells(writeRow, 3).Value = key
        writeRow = writeRow + 1
    Next key

    MsgBox "Process terminé avec succès.", vbInformation

End Sub