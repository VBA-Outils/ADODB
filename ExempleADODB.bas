Attribute VB_Name = "ExempleADODB"
Option Explicit

' Exemple d'utilisation de la classe ADODB

Public Sub EcrireLireFichiers()

    ' D�claration d'un objet ADODB
    Dim oFichier As New ADODB
    ' Enregistrement et nom du fichier
    Dim sEnreg As String, sNomFichier As String
    
    ' Cr�er un fichier texte
    
    ' Ouverture du fichier CSV en �criture
    With oFichier
        .TypeFichier = FICHIER_TEXTE
        .TypeAcces = ACCES_ECRITURE
        .Encodage = UTF_8
        .Ouvrir
        ' Ecriture dans le flux ADODB de l'ent�te du fichier texte
        sEnreg = "Marque;Modele;Categorie;Carburant;Puissance" & vbCrLf
        .EcrireEnregistrement (sEnreg)
        ' Ecriture des donnees dans le flux ADODB
        sEnreg = "Marque1;Mod�le1;SUV;Essence;129 ch" & vbCrLf
        .EcrireEnregistrement (sEnreg)
        sEnreg = "Marque2;Mod�le2;SUV;Essence;89 ch" & vbCrLf
        .EcrireEnregistrement (sEnreg)
        ' S�lectionner le nom du fichier
        .NomInitialFichier = Environ("OneDrive") & "\Documents\tests.txt"
        .ExtensionFiltre = "*.txt"
        .TitreBoiteDeDialogue = "Enregistrer le fichier sous"
        .SelectionnerFichierEnregistrerSous
        ' Si le nom du fichier n'a pas �t� s�lectionn� alors on quitte
        If .NomFichier = "" Then Exit Sub
        ' R�cup�re le nom du fichier afin de pouvoir le lire dans l'�tape suivante
        sNomFichier = .NomFichier
        ' Enregistrement du flux ADODB
        .EnregistrerSous
        ' Fermeture du flux
        .Fermer
    End With
    
    ' Lire le fichier texte pr�c�demment cr��
    
    ' Ouverture du fichier en lecture
    With oFichier
        .TypeFichier = FICHIER_TEXTE
        .TypeAcces = ACCES_LECTURE
        .Encodage = UTF_8
        .SeparateurLigne = SEPARATEUR_CRLF
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
    
    ' Lib�rer les ressources
    Set oFichier = Nothing
    
End Sub
