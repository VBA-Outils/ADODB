Attribute VB_Name = "ExempleADODB"
Option Explicit

' Exemple d'utilisation de la classe ADODB

Public Sub EcrireLireFichiers()

    ' Declaration d'un objet
    Dim oFichier As New ADODB
    Dim sEnreg As String, sNomFichier As String
    
    ' Creer un fichier texte
    sNomFichier = Environ("OneDrive") & "\Documents\tests.txt"
    
    ' Ouverture du fichier CSV en ecriture
    With oFichier
        .TypeFichier = FICHIER_TEXTE
        .TypeAcces = ACCES_ECRITURE
        .Encodage = UTF_8
        .NomFichier = sNomFichier
        .Ouvrir
    End With
    ' Ecriture dans le flux ADODB de l'entete du fichier texte
    sEnreg = "Marque;Modele;Categorie;Carburant;Puissance" & vbCrLf
    oFichier.EcrireEnregistrement (sEnreg)
    ' Ecriture des donnees dans le flux ADODB
    sEnreg = "Suzuki;Vitara;SUV;Essence;129 ch" & vbCrLf
    oFichier.EcrireEnregistrement (sEnreg)
    sEnreg = "Suzuki;Swift;SUV;Essence;89 ch" & vbCrLf
    oFichier.EcrireEnregistrement (sEnreg)
    ' Enregistrement du flux ADODB
    oFichier.EnregistrerSous
    ' Fermeture du flux
    oFichier.Fermer
    
    ' Lire le fichier texte precedemment cree
    
    ' Ouverture du fichier en lecture
    With oFichier
        .TypeFichier = FICHIER_TEXTE
        .TypeAcces = ACCES_LECTURE
        .Encodage = UTF_8
        .SeparateurLigne = SEPARATEUR_CRLF
        .NomFichier = sNomFichier
        .Ouvrir
    End With
    
    ' Lecture des enregistrements
    While Not oFichier.FinFichier
        sEnreg = oFichier.LireEnregistrement
        ' Afficher l'enregistrement lu dans la console VBA
        Debug.Print sEnreg
    Wend
    
    ' Fermeture du fichier
    oFichier.Fermer
    
    ' Liberer les ressources
    Set oFichier = Nothing
    
End Sub
