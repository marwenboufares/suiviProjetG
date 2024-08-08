Sub GenerateFilesFromTemplate()
    Dim templatePath As String
    Dim outputDir As String
    Dim fileInfos As Variant
    Dim wb As Workbook
    Dim newWb As Workbook
    Dim ws As Worksheet
    Dim i As Integer

    ' Définissez le chemin vers le modèle
    templatePath = "C:\Users\mboufares\Desktop\suiviProjetG\BDD\COLLAB.xlsx" ' Remplacez par le chemin réel vers votre modèle

    ' Définissez le répertoire de sortie
    outputDir = "C:\Users\mboufares\Desktop\suiviProjetG\generatedProject\" ' Remplacez par le chemin réel vers le dossier de sortie
    
    ' Créez le répertoire de sortie s'il n'existe pas
    If Dir(outputDir, vbDirectory) = "" Then
        MkDir outputDir
    End If
    
    ' Définissez les informations des fichiers (nom du fichier, info)
    ' Prenez les noms des fichiers des cellules D5 à D7
    Set ws = ThisWorkbook.Sheets("overview") ' Remplacez "overview" par le nom de votre feuille de calcul
    fileInfos = Array( _
        Array(ws.Range("D5").Value, "Collaborateur_1"), _
        Array(ws.Range("D6").Value, "Collaborateur_2"), _
        Array(ws.Range("D7").Value, "Collaborateur_3") _
    )
    
    ' Récupérez le nom du projet de la cellule B5
    Dim projetNom As String
    projetNom = ws.Range("B5").Value
    
    ' Ouvrez le modèle
    Set wb = Workbooks.Open(templatePath)
    
    ' Boucle à travers chaque entrée dans fileInfos
    For i = LBound(fileInfos) To UBound(fileInfos)
        ' Créez un nouveau fichier basé sur le modèle
        wb.Sheets.Copy
        Set newWb = ActiveWorkbook
        
        ' Vérifiez que newWb est bien défini
        If Not newWb Is Nothing Then
            ' Modifiez les cellules nécessaires dans le nouveau fichier
            With newWb.Sheets(1)
                .Range("B6").Value = fileInfos(i)(0) ' Nom
            End With
            
            ' Collez le nom du projet dans la cellule A3 de la feuille "BDD"
            On Error Resume Next ' En cas d'erreur, passez à l'instruction suivante
            newWb.Sheets("BDD").Range("A3").Value = projetNom
            On Error GoTo 0 ' Réactivez la gestion des erreurs
            
            ' Sauvegardez le nouveau fichier
            newWb.SaveAs Filename:=outputDir & fileInfos(i)(0) & ".xlsx"
            newWb.Close
        Else
            MsgBox "Erreur lors de la copie de la feuille pour " & fileInfos(i)(0), vbCritical
        End If
    Next i
    
    ' Fermez le modèle sans sauvegarder les modifications
    wb.Close SaveChanges:=False
    
    MsgBox "Les fichiers ont été générés avec succès dans " & outputDir
End Sub
