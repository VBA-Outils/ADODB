Attribute VB_Name = "ExempleADODB"
Option Explicit

' Exemple d'utilisation de la classe ADODB

Public Sub EcrireLireFichiers()

    ' Déclaration d'un objet ADODB
    Dim oFichier As New ADODB
    ' Enregistrement et nom du fichier
    Dim sEnreg As String, sNomFichier As String
    
    ' Créer un fichier texte
    
    ' Ouverture du fichier CSV en écriture
    With oFichier
        .TypeFichier = AD_TYPE_TEXT
        .TypeAcces = AD_MODE_WRITE
        .Encodage = AD_UTF_8
        .Ouvrir
        ' Ecriture dans le flux ADODB de l'entête du fichier texte
        sEnreg = "Marque;Modele;Categorie;Carburant;Puissance" & vbCrLf
        .EcrireEnregistrement (sEnreg)
        ' Ecriture des donnees dans le flux ADODB
        sEnreg = "Marque1;Modèle1;SUV;Essence;129 ch" & vbCrLf
        .EcrireEnregistrement (sEnreg)
        sEnreg = "Marque2;Modèle2;SUV;Essence;89 ch" & vbCrLf
        .EcrireEnregistrement (sEnreg)
        ' Sélectionner le nom du fichier
        .NomInitialFichier = Environ("OneDrive") & "\Documents\tests.txt"
        .ExtensionFiltre = "*.txt"
        .TitreBoiteDeDialogue = "Enregistrer le fichier sous"
        .SelectionnerFichierEnregistrerSous
        ' Si le nom du fichier n'a pas été sélectionné alors on quitte
        If .NomFichier = "" Then Exit Sub
        ' Récupère le nom du fichier afin de pouvoir le lire dans l'étape suivante
        sNomFichier = .NomFichier
        ' Enregistrement du flux ADODB
        .EnregistrerSous
        ' Fermeture du flux
        .Fermer
    End With
    
    ' Lire le fichier texte précédemment créé
    
    ' Ouverture du fichier en lecture
    With oFichier
        .TypeFichier = AD_TYPE_TEXT
        .TypeAcces = AD_MODE_READ
        .Encodage = AD_UTF_8
        .SeparateurLigne = AD_CR_LF
        .NomFichier = sNomFichier
        .Ouvrir
    End With
    
    ' Lecture des enregsitrements
    While Not oFichier.FinFichier
        sEnreg = oFichier.LireEnregistrement
        ' Afficher l'enregsitrement lu dans la console VBA
        Debug.Print sEnreg
    Wend
    
    ' Fermeture du fichier
    oFichier.Fermer
    
    ' Libérer les ressources
    Set oFichier = Nothing
    
End Sub
