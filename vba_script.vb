Option Explicit

' Numéro 1 : traitement des cours manquants

Sub RemplirCoursManquants()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    For Each ws In ThisWorkbook.Sheets
        lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        
        For i = 2 To lastRow - 1
            If IsEmpty(ws.Cells(i, 2)) Then
                ws.Cells(i, 2).Value = (ws.Cells(i - 1, 2).Value + ws.Cells(i + 1, 2).Value) / 2
                ws.Cells(i, 2).Font.Color = RGB(255, 0, 0)
                ws.Cells(i, 2).Font.Bold = True
                
                ws.Cells(i, 4).Value = "Cours hypothétique"
                ws.Cells(i, 4).Font.Color = RGB(255, 0, 0)
                ws.Cells(i, 4).Font.Bold = True
            End If
        Next i
    Next ws
    
    MsgBox "Les cours manquants ont été remplis avec succès !"
End Sub

'Numéro 2 : Calculons les rentabilité journalières

Sub CalculerRentabilitesJournalieres()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    For Each ws In ThisWorkbook.Sheets
        lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        
        If lastRow > 1 Then
            ws.Cells(1, 3).Value = "Rentabilité J"
            
            For i = 2 To lastRow - 1
                If ws.Cells(i + 1, 2).Value <> 0 And ws.Cells(i, 2).Value <> 0 Then
                    ws.Cells(i, 3).Value = WorksheetFunction.Ln(ws.Cells(i + 1, 2).Value / ws.Cells(i, 2).Value)
                    ws.Cells(i, 3).NumberFormat = "0.00%"
                    ws.Cells(i, 3).Font.Color = RGB(128, 0, 128)
                    ws.Cells(i, 3).Font.Bold = True
                Else
                    ws.Cells(i, 3).Value = "Erreur"
                End If
            Next i
        End If
    Next ws
    
    MsgBox "Les rentabilités journalières ont été calculées avec succès !"
End Sub

'Numéro 3 : rentabilité annuelle

Sub CalculerRentabilitesAnnuelles()
    Dim ws As Worksheet, wsResult As Worksheet
    Dim lastRow As Long, startRow As Long, endRow As Long
    Dim i As Long
    Dim startDate As Date, endDate As Date
    Dim resultRow As Long
    Dim yearStart As String, yearEnd As String
    Dim rendementAnnuel As Double
    
    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("Rentabilités annuelles")
    If wsResult Is Nothing Then
        Set wsResult = ThisWorkbook.Sheets.Add
        wsResult.Name = "Rentabilités annuelles"
    Else
        wsResult.Cells.Clear
    End If
    On Error GoTo 0
    
    wsResult.Cells(1, 1).Value = "Année"
    wsResult.Cells(1, 2).Value = "Crypto-monnaie"
    wsResult.Cells(1, 3).Value = "Rentabilité annuelle"
    wsResult.Columns(3).NumberFormat = "0.00%"
    
    resultRow = 2
    startDate = DateSerial(2019, 1, 1)
    endDate = DateSerial(2024, 12, 31)
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Rentabilités annuelles" Then
            lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
            
            For i = 2 To lastRow - 1
                If ws.Cells(i, 1).Value >= startDate And ws.Cells(i, 1).Value < endDate Then
                    startRow = i
                    Do While ws.Cells(i, 1).Value < DateAdd("yyyy", 1, ws.Cells(startRow, 1).Value) And i <= lastRow
                        i = i + 1
                    Loop
                    endRow = i - 1
                    
                    
                    If ws.Cells(startRow, 2).Value > 0 And ws.Cells(endRow, 2).Value > 0 Then
                        rendementAnnuel = (ws.Cells(endRow, 2).Value / ws.Cells(startRow, 2).Value - 1)
                        yearStart = Year(ws.Cells(startRow, 1).Value)
                        yearEnd = Year(ws.Cells(endRow, 1).Value)
                        
                        wsResult.Cells(resultRow, 1).Value = yearStart & "-" & yearEnd
                        wsResult.Cells(resultRow, 2).Value = ws.Name
                        wsResult.Cells(resultRow, 3).Value = rendementAnnuel
                        wsResult.Cells(resultRow, 3).Font.Color = RGB(128, 0, 128)
                        wsResult.Cells(resultRow, 3).Font.Bold = True
                        
                        resultRow = resultRow + 1
                    End If
                End If
            Next i
        End If
    Next ws
    
    MsgBox "Les rentabilités annuelles ont été calculées avec succès !"
End Sub

' Numéro 4 : Calcul de la rentabilité annualisé

Sub CalculerRentabiliteAnnualisee()
    Dim ws As Worksheet
    Dim wsResult As Worksheet
    Dim lastRow As Long
    Dim rendementCumulé As Double
    Dim rendementAnnualisé As Double
    Dim prixDebut As Double, prixFin As Double
    Dim i As Long
    Dim n As Double
    Dim resultRow As Long
    

    n = 5

    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("Rentabilités annuelles")
    If wsResult Is Nothing Then
        MsgBox "La feuille 'Rentabilités annuelles' n'existe pas. Veuillez la créer avant d'exécuter la macro.", vbExclamation, "Erreur"
        Exit Sub
    End If
    On Error GoTo 0


    resultRow = wsResult.Cells(wsResult.Rows.Count, 5).End(xlUp).Row + 1

    If resultRow = 2 Then
        wsResult.Cells(1, 5).Value = "Crypto-monnaie"
        wsResult.Cells(1, 6).Value = "Rentabilité Cumulée"
        wsResult.Cells(1, 7).Value = "Rentabilité Annualisée"
        wsResult.Rows(1).Font.Bold = True
    End If


    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsResult.Name Then
    
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            prixDebut = 0
            prixFin = 0
            
            For i = 2 To lastRow
                If ws.Cells(i, 1).Value = DateSerial(2019, 12, 1) Then
                    prixDebut = ws.Cells(i, 2).Value
                End If
                If ws.Cells(i, 1).Value = DateSerial(2024, 12, 1) Then
                    prixFin = ws.Cells(i, 2).Value
                End If
            Next i

        
            If prixDebut > 0 And prixFin > 0 Then
        
                rendementCumulé = (prixFin / prixDebut) - 1

            
                rendementAnnualisé = (1 + rendementCumulé) ^ (1 / n) - 1

            
                wsResult.Cells(resultRow, 5).Value = ws.Name
                wsResult.Cells(resultRow, 6).Value = rendementCumulé
                wsResult.Cells(resultRow, 7).Value = rendementAnnualisé

            
                wsResult.Cells(resultRow, 6).NumberFormat = "0.00%"
                wsResult.Cells(resultRow, 7).NumberFormat = "0.00%"
                
    
                wsResult.Cells(resultRow, 6).Font.Bold = True
                wsResult.Cells(resultRow, 6).Font.Color = RGB(128, 0, 128)
                wsResult.Cells(resultRow, 7).Font.Bold = True
                wsResult.Cells(resultRow, 7).Font.Color = RGB(128, 0, 128)
                
                resultRow = resultRow + 1
            End If
        End If
    Next ws

    MsgBox "Les rentabilités annuelles ont été calculées avec succès et ajoutées à la feuille 'Rentabilités annuelles' !", vbInformation, "Succès"
End Sub

'Numéro 5 : volatilité annualisée

Sub AjouterVolatiliteAnnualisee()
    Dim ws As Worksheet
    Dim wsResult As Worksheet
    Dim lastRow As Long
    Dim rendementCol As Range
    Dim rendements() As Double
    Dim i As Long
    Dim sigmaJournaliere As Double
    Dim sigmaAnnualisee As Double

    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("Rentabilités annuelles")
    If wsResult Is Nothing Then
        MsgBox "La feuille 'Rentabilités annuelles' n'existe pas. Veuillez la créer avant d'exécuter la macro.", vbExclamation, "Erreur"
        Exit Sub
    End If
    On Error GoTo 0

    lastRow = wsResult.Cells(wsResult.Rows.Count, 5).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "Aucune donnée trouvée dans la colonne des crypto-monnaies. Veuillez d'abord compléter les colonnes E, F et G.", vbExclamation, "Erreur"
        Exit Sub
    End If

    wsResult.Cells(1, 8).Value = "Volatilité Annualisée"
    wsResult.Cells(1, 8).Font.Bold = True

    For i = 2 To lastRow
        Dim cryptoName As String
        cryptoName = wsResult.Cells(i, 5).Value

        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(cryptoName)
        On Error GoTo 0

        If Not ws Is Nothing Then
            lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row

            Set rendementCol = ws.Range("C2:C" & lastRow)
            
            ReDim rendements(1 To rendementCol.Rows.Count)
            Dim j As Long
            For j = 1 To rendementCol.Rows.Count
                rendements(j) = rendementCol.Cells(j, 1).Value
            Next j

            sigmaJournaliere = WorksheetFunction.StDev(rendements)
            sigmaAnnualisee = sigmaJournaliere * Sqr(365)
            wsResult.Cells(i, 8).Value = sigmaAnnualisee

            wsResult.Cells(i, 8).NumberFormat = "0.00%"

            wsResult.Cells(i, 8).Font.Bold = True
            wsResult.Cells(i, 8).Font.Color = RGB(255, 69, 0)
        Else
            wsResult.Cells(i, 8).Value = "Feuille manquante"
            wsResult.Cells(i, 8).Font.Color = RGB(255, 0, 0)
            wsResult.Cells(i, 8).Font.Bold = True
        End If
    Next i

    MsgBox "Les volatilités annualisées ont été ajoutées dans la colonne H.", vbInformation, "Succès"
End Sub



'Numéro 6 : Sharpe Ratio

Sub CalculerSharpeRatio()
    Dim wsResult As Worksheet
    Dim lastRow As Long
    Dim rendementAnnualise As Double
    Dim volatiliteAnnualisee As Double
    Dim ratioSharpe As Double
    Dim tauxSansRisque As Double
    Dim i As Long

    ' Définir le taux sans risque
    tauxSansRisque = 0.002


    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("Rentabilités annuelles")
    If wsResult Is Nothing Then
        MsgBox "La feuille 'Rentabilités annuelles' n'existe pas. Veuillez la créer avant d'exécuter la macro.", vbExclamation, "Erreur"
        Exit Sub
    End If
    On Error GoTo 0

    lastRow = wsResult.Cells(wsResult.Rows.Count, 5).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "Aucune donnée trouvée dans la colonne des crypto-monnaies. Veuillez d'abord compléter les colonnes E, F, G et H.", vbExclamation, "Erreur"
        Exit Sub
    End If

    wsResult.Cells(1, 9).Value = "Sharpe Ratio"
    wsResult.Cells(1, 9).Font.Bold = True

    For i = 2 To lastRow
        rendementAnnualise = wsResult.Cells(i, 7).Value
        volatiliteAnnualisee = wsResult.Cells(i, 8).Value

        If volatiliteAnnualisee > 0 Then
            ratioSharpe = (rendementAnnualise - tauxSansRisque) / volatiliteAnnualisee
            wsResult.Cells(i, 9).Value = ratioSharpe
        Else
            wsResult.Cells(i, 9).Value = "N/A"
        End If

        If IsNumeric(wsResult.Cells(i, 9).Value) Then
            wsResult.Cells(i, 9).NumberFormat = "0.00"
            wsResult.Cells(i, 9).Font.Bold = True
            wsResult.Cells(i, 9).Font.Color = RGB(128, 0, 128)
        Else
            wsResult.Cells(i, 9).Font.Color = RGB(255, 0, 0)
            wsResult.Cells(i, 9).Font.Bold = True
        End If
    Next i

    MsgBox "Les ratios de Sharpe ont été calculés et ajoutés dans la colonne I.", vbInformation, "Succès"
End Sub

'Numéro 7 : interpretation intermediaire

Sub AfficherMessagePerformance()
    Dim wsResult As Worksheet
    Dim maxSharpe As Double
    Dim bestCrypto As String
    Dim i As Long
    Dim message As String

    ' Accéder à la feuille "Rentabilités annuelles"
    Set wsResult = ThisWorkbook.Sheets("Rentabilités annuelles")

    ' Initialiser les variables
    maxSharpe = -9999
    bestCrypto = ""

    For i = 2 To wsResult.Cells(wsResult.Rows.Count, 9).End(xlUp).Row
        If IsNumeric(wsResult.Cells(i, 9).Value) And wsResult.Cells(i, 9).Value > maxSharpe Then
            maxSharpe = wsResult.Cells(i, 9).Value
            bestCrypto = wsResult.Cells(i, 5).Value
        End If
    Next i

    If bestCrypto <> "" Then
        message = "La crypto-monnaie ayant le ratio de Sharpe le plus élevé est " & bestCrypto & _
                  " avec un ratio de " & Format(maxSharpe, "0.00") & ". Cela signifie que sur la période de 2019 à 2024, " & _
                  bestCrypto & " a offert la meilleure performance ajustée au risque parmi les trois crypto-monnaies étudiées."
    Else
        message = "Aucune crypto-monnaie valide trouvée pour le ratio de Sharpe."
    End If

    
    With wsResult.Range("F19:M21")
        .Merge
        .Value = message
        .Font.Bold = True
        .Font.Color = RGB(128, 0, 128)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    wsResult.Rows("19:20").AutoFit
End Sub

' Numéro 8:Portefeuille équipondéré

'ETAPE 1
Sub CalculerPortefeuilleEquipondere()
    Dim ws As Worksheet
    Dim portfolioWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim BTCWs As Worksheet, ETHWs As Worksheet, XRPWs As Worksheet
    
    ' Définir les feuilles pour chaque crypto-monnaie
    Set BTCWs = ThisWorkbook.Sheets("BTC-USD")
    Set ETHWs = ThisWorkbook.Sheets("ETH-USD")
    Set XRPWs = ThisWorkbook.Sheets("XRP-USD")
    
    Set portfolioWs = ThisWorkbook.Worksheets.Add
    portfolioWs.Name = "Portefeuille Équipondéré"
    
    portfolioWs.Cells(1, 1).Value = "Date"
    portfolioWs.Cells(1, 2).Value = "Cours Portefeuille"
    portfolioWs.Rows(1).Font.Bold = True
    
    lastRow = BTCWs.Cells(BTCWs.Rows.Count, "A").End(xlUp).Row
    BTCWs.Range("A2:A" & lastRow).Copy Destination:=portfolioWs.Range("A2")
    
    For i = 2 To lastRow
        portfolioWs.Cells(i, 2).Value = (BTCWs.Cells(i, 2).Value + ETHWs.Cells(i, 2).Value + XRPWs.Cells(i, 2).Value) / 3
    Next i
End Sub

'Etape 2

Sub CalculerRentabilitesJournalieres_equiponderee()
    Dim wsPortefeuille As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set wsPortefeuille = ThisWorkbook.Sheets("Portefeuille Équipondéré")

    lastRow = wsPortefeuille.Cells(wsPortefeuille.Rows.Count, 1).End(xlUp).Row

    wsPortefeuille.Cells(1, 3).Value = "Rentabilité Journalière"
    wsPortefeuille.Rows(1).Font.Bold = True

    For i = 2 To lastRow - 1 '
        wsPortefeuille.Cells(i, 3).Value = WorksheetFunction.Ln(wsPortefeuille.Cells(i + 1, 2).Value / wsPortefeuille.Cells(i, 2).Value)
    Next i

    wsPortefeuille.Cells(lastRow, 3).Value = "N/A"

    With wsPortefeuille.Columns(3)
        .NumberFormat = "0.00%"
        .Font.Color = RGB(128, 0, 128)
        .Font.Bold = True
    End With

    ' Message de confirmation
    MsgBox "Les rentabilités journalières ont été calculées et ajoutées à la feuille 'Portefeuille Équipondéré'.", vbInformation
End Sub


'Etape 3

Sub CalculerRentabilitesAnnuelles_equipon()
    Dim wsPortefeuille As Worksheet, wsRentabilites As Worksheet
    Dim lastRowPortefeuille As Long, lastRowRentabilites As Long
    Dim startRow As Long, endRow As Long
    Dim startDate As Date, endDate As Date
    Dim rendementAnnuel As Double
    Dim i As Long

    ' Associer les feuilles
    Set wsPortefeuille = ThisWorkbook.Sheets("Portefeuille Équipondéré")
    Set wsRentabilites = ThisWorkbook.Sheets("Rentabilités annuelles")

    ' Identifier les dernières lignes
    lastRowPortefeuille = wsPortefeuille.Cells(wsPortefeuille.Rows.Count, 1).End(xlUp).Row
    lastRowRentabilites = wsRentabilites.Cells(wsRentabilites.Rows.Count, 1).End(xlUp).Row

    ' Ajouter les en-têtes si la feuille est vide
    If lastRowRentabilites = 1 Then
        With wsRentabilites
            .Cells(1, 1).Value = "Année"
            .Cells(1, 2).Value = "Crypto-monnaie"
            .Cells(1, 3).Value = "Rentabilité annuelle"
            .Rows(1).Font.Bold = True
            .Rows(1).Interior.Color = RGB(200, 230, 255) ' Bleu clair
        End With
    End If

    
    For i = 2 To lastRowPortefeuille - 1
        If wsPortefeuille.Cells(i, 1).Value >= DateSerial(2019, 12, 1) Then
            startRow = i
            Do While wsPortefeuille.Cells(i, 1).Value < DateAdd("yyyy", 1, wsPortefeuille.Cells(startRow, 1).Value) And i <= lastRowPortefeuille
                i = i + 1
            Loop
            endRow = i - 1

            rendementAnnuel = (wsPortefeuille.Cells(endRow, 2).Value / wsPortefeuille.Cells(startRow, 2).Value) - 1

            With wsRentabilites
                .Cells(lastRowRentabilites + 1, 1).Value = Year(wsPortefeuille.Cells(startRow, 1).Value) & "-" & Year(wsPortefeuille.Cells(endRow, 1).Value)
                .Cells(lastRowRentabilites + 1, 2).Value = "Portefeuille Équipondéré"
                .Cells(lastRowRentabilites + 1, 3).Value = Format(rendementAnnuel * 100, "0.00") & "%"
                .Cells(lastRowRentabilites + 1, 3).Font.Color = RGB(0, 128, 0)
                .Cells(lastRowRentabilites + 1, 3).Font.Bold = True
            End With

            lastRowRentabilites = lastRowRentabilites + 1
        End If
    Next i

End Sub

'Etape 4
Sub CalculerRentabiliteCumulee_equipon()
    Dim wsPortefeuille As Worksheet, wsRentabilites As Worksheet
    Dim lastRowPortefeuille As Long
    Dim rendementCumule As Double

    Set wsPortefeuille = ThisWorkbook.Sheets("Portefeuille Équipondéré")
    Set wsRentabilites = ThisWorkbook.Sheets("Rentabilités annuelles")

    lastRowPortefeuille = wsPortefeuille.Cells(wsPortefeuille.Rows.Count, 1).End(xlUp).Row

    rendementCumule = (wsPortefeuille.Cells(lastRowPortefeuille, 2).Value / wsPortefeuille.Cells(2, 2).Value) - 1

    With wsRentabilites
        .Cells(5, 5).Value = "Portefeuille Équipondéré"
        .Cells(5, 6).Value = Format(rendementCumule * 100, "0.00") & "%"
        .Cells(5, 6).Font.Color = RGB(128, 0, 128)
        .Cells(5, 6).Font.Bold = True
    End With

End Sub

'Etape 5
Sub CalculerRentabiliteAnnualisee_equipon()
    Dim wsPortefeuille As Worksheet, wsRentabilites As Worksheet
    Dim lastRowPortefeuille As Long
    Dim rendementCumule As Double, rendementAnnualise As Double
    Dim nbAnnees As Double

    Set wsPortefeuille = ThisWorkbook.Sheets("Portefeuille Équipondéré")
    Set wsRentabilites = ThisWorkbook.Sheets("Rentabilités annuelles")

    lastRowPortefeuille = wsPortefeuille.Cells(wsPortefeuille.Rows.Count, 1).End(xlUp).Row

    nbAnnees = 5

    rendementCumule = (wsPortefeuille.Cells(lastRowPortefeuille, 2).Value / wsPortefeuille.Cells(2, 2).Value) - 1

    rendementAnnualise = (1 + rendementCumule) ^ (1 / nbAnnees) - 1

    With wsRentabilites
        .Cells(5, 6).Value = Format(rendementCumule * 100, "0.00") & "%" '
        .Cells(5, 6).Font.Color = RGB(128, 0, 128)
        .Cells(5, 6).Font.Bold = True

        .Cells(5, 7).Value = Format(rendementAnnualise * 100, "0.00") & "%" '
        .Cells(5, 7).Font.Color = RGB(128, 0, 128)
        .Cells(5, 7).Font.Bold = True
    End With

    ' Confirmation
    MsgBox "Les rentabilités cumulée et annualisée ont été ajoutées avec succès.", vbInformation
End Sub


'Etape 6 : volatilité annualisée equipon

Sub CalculerVolatiliteAnnualisee_Portefeuille()
    Dim wsPortefeuille As Worksheet
    Dim wsResultat As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim rentabilites() As Double
    Dim j As Long
    Dim ecartTypeJournalier As Double
    Dim volatiliteAnnualisee As Double

    Set wsPortefeuille = ThisWorkbook.Sheets("Portefeuille équipondéré")
    Set wsResultat = ThisWorkbook.Sheets("Rentabilités annuelles")

    lastRow = wsPortefeuille.Cells(wsPortefeuille.Rows.Count, 3).End(xlUp).Row
    ReDim rentabilites(1 To lastRow - 1)

    j = 1
    For i = 2 To lastRow
        If IsNumeric(wsPortefeuille.Cells(i, 3).Value) Then
            rentabilites(j) = wsPortefeuille.Cells(i, 3).Value
            j = j + 1
        End If
    Next i

    ecartTypeJournalier = Application.WorksheetFunction.StDev(rentabilites)

    volatiliteAnnualisee = ecartTypeJournalier * Sqr(365)

    wsResultat.Cells(5, 8).Value = volatiliteAnnualisee
    wsResultat.Range("H5").Font.Bold = True
    wsResultat.Cells(5, 8).Font.Color = RGB(255, 0, 0)
    wsResultat.Cells(5, 8).NumberFormat = "0.00%"

End Sub

'Etape 7 : ratio sharpe equipon

Sub CalculerSharpeRatio_Portefeuille()
    Dim wsResultat As Worksheet
    Dim rentabiliteAnnualisee As Double
    Dim volatiliteAnnualisee As Double
    Dim tauxSansRisque As Double
    Dim sharpeRatio As Double

    tauxSansRisque = 0.002

    Set wsResultat = ThisWorkbook.Sheets("Rentabilités annuelles")

    rentabiliteAnnualisee = wsResultat.Cells(5, 7).Value
    volatiliteAnnualisee = wsResultat.Cells(5, 8).Value

    If volatiliteAnnualisee = 0 Then
        MsgBox "La volatilité annualisée est nulle. Impossible de calculer le ratio de Sharpe.", vbExclamation, "Erreur"
        Exit Sub
    End If

    sharpeRatio = (rentabiliteAnnualisee - tauxSansRisque) / volatiliteAnnualisee

    wsResultat.Cells(5, 9).Value = sharpeRatio
    wsResultat.Cells(5, 9).NumberFormat = "0.00"
    wsResultat.Cells(5, 9).Font.Bold = True
    wsResultat.Cells(5, 9).Font.Color = RGB(128, 0, 128)
End Sub


'Numéro 9 : interpretation finale
Sub ComparerSharpeRatio()
    Dim wsResultat As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim maxSharpe As Double
    Dim bestAsset As String
    Dim currentSharpe As Double

    
    Set wsResultat = ThisWorkbook.Sheets("Rentabilités annuelles")

    lastRow = wsResultat.Cells(wsResultat.Rows.Count, 5).End(xlUp).Row

    maxSharpe = -9999
    bestAsset = ""

    For i = 2 To lastRow
        If IsNumeric(wsResultat.Cells(i, 9).Value) Then
            currentSharpe = wsResultat.Cells(i, 9).Value

            If currentSharpe > maxSharpe Then
                maxSharpe = currentSharpe
                bestAsset = wsResultat.Cells(i, 5).Value
            End If
        End If
    Next i

    If bestAsset <> "" Then
        
        With wsResultat.Range("E8:M13")
            .Merge
            .Value = "Le ratio de Sharpe du " & bestAsset & " est de " & Format(maxSharpe, "0.00") & "." & vbNewLine & _
                     "Cela signifie que, sur la période 2019 à 2024, " & bestAsset & " a offert une meilleure performance ajustée au risque " & _
                     "par rapport aux trois cryptoactifs individuels pris séparément. En d'autres termes, " & bestAsset & " a généré " & _
                     "un meilleur rendement par unité de risque pris, comparé à un investissement uniquement dans Bitcoin, Ethereum ou XRP."
            .Font.Bold = True
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(230, 230, 250)
            .Font.Color = RGB(0, 0, 139) '
        End With

        MsgBox "Le ratio de Sharpe le plus élevé a été affiché avec succès dans la plage F19:H23.", vbInformation, "Succès"
    Else
        MsgBox "Aucun ratio de Sharpe valide n'a été trouvé.", vbExclamation, "Erreur"
    End If
End Sub




