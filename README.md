# Outils_VBA
 
Exemples d'utilisation de la classe ADODB


    Dim oFichier As ADODB, sEnreg As String
    
    ' Lecture d'un enregistrement d'un fichier
    Set oFichier = New ADODB
    With oFichier
        .Charset = "utf-8"
        .TypeFichier = TYPE_FICHIER_TEXTE
        .NomFichier = ENVIRON("OneDrive") & "\Documents\test.json"
        .TypeAcces = ACCES_LECTURE
        .Ouvrir
        .SeparateurLigne = SEPARATEUR_CRLF
        sEnreg = .LireFichier
        .Fermer
    End With
    Set oFichier = Nothing

    ' Ecriture d'un fichier csv
    Set oFichier = New ADODB
    With oFichier
        .TypeFichier = TYPE_FICHIER_TEXTE
        .TitreBoiteDeDialogue = "Nom du fichier"
        .Filtre = INDEX_FILTRE_CSV_TAB
        .SelectionnerFichierEnregistrerSous
        .TypeAcces = ACCES_ECRITURE
        .Encodage = UTF_8
        .SeparateurLigne = SEPARATEUR_CRLF
        .Ouvrir
        .EcrireEnregistrement (sEnreg)
        .EnregistrerSous
        .Fermer
    End With
    Set oFichier = Nothing

    ' Lecture complète d'un fichier
    Set oFichier = New ADODB
    With oFichier
        .TypeFichier = TYPE_FICHIER_BINAIRE
        .NomFichier = "C:\Office 2021.png"
        .TypeAcces = ACCES_LECTURE
        .Ouvrir
        sEnreg = .LireFichier
        .Fermer
    End With
    Set oFichier = Nothing
