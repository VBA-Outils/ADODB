VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADODB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' https://github.com/VBA-Outils/ADODB
'
' Gestion des fichiers via ADODB
'
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Copyright (c) 2024, Vincent ROSSET
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Erreurs:
'  1000   : Erreur lors de l'appel d'une propriété ou méthode. Voir le message d'erreur.

' Nécessite d'activer la référence "Microsoft Scripting RunTime"
'
' Dans l'éditeur de macros (Alt+F11): Menu Outils \ Références
' Cochez les lignes :
' -- "Microsoft Scripting RunTime"
' Cliquez sur le bouton OK pour valider.

Option Explicit

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'
' Propriétés de la classe :
' -------------------------
'   TypeFichier                        : Type du fichier (texte ou binaire)
'   Encodage                           : Encodage du fichier (enum)
'   EncodageTxt                        : Encodage du fichier (saisie libre)
'   SeparateurLigne                    : Séparateur de lignes d'un fichier texte
'   TypeAcces                          : Accès au fichier (lecture, écriture)
'   TitreBoiteDeDialogue               : Titre affiché dans les boîtes de dialogue
'   LibelleFiltre                      : Libellé du filtre utilisé pour afficher les fichiers dans une boîte de dialogue
'   ExtensionFiltre                    : Filtre des extensions de fichier utilisé dans une boîte de dialogue
'   NomInitialFichier                  : Nom initial du fichier affiché dans une boîte de dialogue lors d'un enregsitrement sous
'   Filtre                             : Filtre des extensions (parmi celles proposées par Excel) de fichier utilisé dans une boîte de dialogue
'   NomFichier                         : nom du fichier (y compris le chemin d'accès)
'   Repertoire                         : Répertoire sélectionné dans une boîte de dialogue
'   Fichier                            : Permettre de pointer sur un objet ADODB Stream en dehors de la classe
'   NbreEnregLus                       : Nombre d'enregistrements lus
'   NbreEnregEcrits                    : Nombre d'enregistrements écrits
'   NbreOctectsLus                     : Nombre d'octets lus
'   NbreOctectsEcrits                  : Nombre d'octets écrits
'
' Méthodes de la classe :
' -----------------------
'   Ouvrir                             : Ouvrir un fichier via ADODB
'   LireFichier                        : Lire l'intégralité d'un fichier (texte ou binaire)
'   LireEnregistrement                 : Lire un fichier texte jusqu'au prochain séparateur
'   FinFichier                         : Fin du fichier texte atteinte
'   Lire                               : Lire n caractères d'un fichier (texte ou binaire)
'   LireEnregistrement                 : Lire le prochain enregistrement (jusqu'au prochain séparateur de lignes) d'un fichier texte
'   LireFichier                        : Lire l'intégralité d'un fichier texte ou binaire
'   Fermer                             : Fermer un fichier
'   Ecrire                             : Ecrire des données dans un fichier texte ou binaire
'   EcrireEnregistrement               : Ecrire un enregistrement dans un fichier texte
'   EnregistrerSous                    : Enregistrer sous le nom de fichier le contenu de l'objet ADODB.Stream
'   SelectionnerFichierEnregistrerSous : Afficher la boîte de dialogue de sélection d'un fichier à enregsitrer sous (avec saisie du nom du fichier)
'   SelectionnerFichier                : Afficher la boîte de dialogue de sélection d'un fichier
'   SelectionnerRepertoire             : Afficher la boîte de dialogue de sélection d'un répertoire
'   EstFichierVolumineux               : Vérifier si un fichier dépasse 4 Go (taille maximale des fichiers texte via ADODB.stream)
'   FichierExiste                      : Vérifier si le fichier dont le nom est "NomFichier" existe
'   RepertoireExiste                   : Vérifier si le répertoire dont le nom est "NomRépertoire" existe
'   LongueurFichier                    : Retourner la longueur d'un fichier en octets
'   Extension                          : Retourner l'extension du fichier (en minuscules)
'
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

Private Enum EnumNonDefini
    NON_DEFINI = 0
End Enum

' Statut du flux
Private Enum EnumStatutFichier
    FICHIER_OUVERT = 1
    FICHIER_FERME = 2
End Enum

' *---------------------------------------------------------------------------------------------------*
' * Propriétés des fichiers                                                                           *
' *---------------------------------------------------------------------------------------------------*

' Encodage du fichier texte
Private pEncodage As String
Private pEncodageNum As CharsetEnum
Private DicEncodage As Dictionary
' Séparateur de lignes (fichier texte uniquement)
Private pSeparateurLigne As LineSeparatorsEnum
' Type de fichier : Texte ou Binaire
Private pTypeFichier As StreamTypeEnum
' Nom physique du fichier (le répertoire peut être inclus dans le nom)
Private pNomFichier As String
' Objet ADODB.stream
Private pFichierADODB As Object
' Type d'accès au fichier : lecture, écriture
Private pTypeAcces As ConnectModeEnum

' *---------------------------------------------------------------------------------------------------*
' * Gestion des boîtes de dialogue                                                                    *
' *---------------------------------------------------------------------------------------------------*

' Titre de la boite de dialogue (fichiers)
Private pTitreBoiteDialogue As String
' Libellé du filtre sur les fichiers à appliquer
Private pLibelleFiltre As String
' Extensions appliquées dans le filtre
Private pExtensionFiltre As String
' Nom du fichier initial proposé dans la boite de dialogue Enregistrer sous
Private pNomInitialFichier As String
' Répertoire sélectionné via la boite de dialogue
Private pRepertoire As String
' Filtre pour la boîte à dialogue d'enregsitrement sous
Private pIndexFiltre As FilterIndexEnum

' *---------------------------------------------------------------------------------------------------*
' * Gestion des fichiers                                                                              *
' *---------------------------------------------------------------------------------------------------*

' Fichier ouvert
Private pStatutFichier As EnumStatutFichier
Private pNbreEnregLus As Long
Private pNbreEnregEcrits As Long
Private pNbreOctectsLus As Long
Private pNbreOctectsEcrits As Long

' *---------------------------------------------------------------------------------------------------*
' * Propriété : objet fichier                                                                         *
' *---------------------------------------------------------------------------------------------------*
Property Set Fichier(FichierAdodb As Object)
    Set pFichierADODB = FichierAdodb
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : type de fichier (texte, binaire)                                                      *
' *---------------------------------------------------------------------------------------------------*
Property Let TypeFichier(TypeFichier As StreamTypeEnum)

    ' Vérifier le type de fichier
    Select Case TypeFichier
        Case AD_TYPE_BINARY
            Select Case True
                Case pEncodage <> ""
                    ' L'encodage n'est pas nécessaire pour lire un fichier binaire
                    Err.Raise -1000, , "L'encodage ne doit pas être défini pour les fichiers binaires."
                Case pSeparateurLigne <> NON_DEFINI
                    ' Le séparateur de lignes n'est pas nécessaire pour lire un fichier binaire
                    Err.Raise -1001, , "Le séparateur de lignes ne doit pas être défini pour les fichiers binaires."
                Case Else
                    pTypeFichier = TypeFichier
            End Select
        Case AD_TYPE_TEXT
            pTypeFichier = TypeFichier
        Case Else
            Err.Raise -1002, , "Le type de fichier " & TypeFichier & " est invalide."
    End Select

End Property

Property Get TypeFichier() As StreamTypeEnum
    TypeFichier = pTypeFichier
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : Encodage du fichier texte (Enum)                                                      *
' *---------------------------------------------------------------------------------------------------*
Property Let Encodage(EncodageNum As CharsetEnum)

    Select Case True
        ' L'encodage doit être renseigné
        Case EncodageNum = NON_DEFINI
            Err.Raise -1010, , "L'encodage n'est pas renseigné."
        Case pTypeFichier = AD_TYPE_BINARY
            ' L'encodage n'est pas nécessaire pour lire un fichier binaire
            Err.Raise -1011, , "L'encodage ne doit pas être défini pour les fichiers binaires."
            ' Vérifie que l'encodage existe dans l'enum
        Case DicEncodage.Exists(EncodageNum)
            pEncodageNum = EncodageNum
            pEncodage = DicEncodage(EncodageNum)
        Case Else
            ' Si l'encodage est absent de l'enum ListeEncodage et du dictionnaire dicEncodage alors une erreur est retournée.
            ' La liste des encodages n'est pas exhaustive et doit être complétée si nécessaire, sinon utiliser la propriété EncodageTxt.
            Err.Raise -1012, , "L'encodage n'existe pas dans l'enum CharsetEnum."
    End Select
    
End Property

Property Get Encodage() As CharsetEnum
    
    Encodage = pEncodageNum
    
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : Encodage du fichier texte (code technique)                                            *
' *---------------------------------------------------------------------------------------------------*
Property Let EncodageTxt(Encodage As String)

    Select Case True
        ' L'encodage doit être renseigné
        Case Trim(Encodage) = ""
            Err.Raise -1020, , "L'encodage n'est pas renseigné."
        ' L'encodage n'est pas nécessaire pour lire un fichier binaire
        Case pTypeFichier = AD_TYPE_BINARY
            Err.Raise -1021, , "L'encodage ne doit pas être défini pour les fichiers binaires."
        ' Aucun contrôle de l'encodage, la saisie est libre.
        Case Else
            pEncodage = Encodage
            pEncodageNum = AD_AUTRE
    End Select
    
End Property

Property Get EncodageTxt() As String
    EncodageTxt = pEncodage
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : séparateur de lignes (CL, LF, CR/LF)                                                  *
' *---------------------------------------------------------------------------------------------------*
Property Let SeparateurLigne(SeparateurLigne As LineSeparatorsEnum)

    If SeparateurLigne = AD_CR Or SeparateurLigne = AD_CR_LF Or SeparateurLigne = AD_LF Then
        If pTypeFichier = AD_TYPE_BINARY Then
            ' Le séparateur ne sert à rien pour les fichiers binaires
            Err.Raise -1030, , "Le séparateur ne doit pas être défini pour les fichiers binaires."
        Else
            pSeparateurLigne = SeparateurLigne
        End If
    Else
        Err.Raise -1031, , "Le séparateur " & SeparateurLigne & " est invalide."
    End If
    
End Property

Property Get SeparateurLigne() As LineSeparatorsEnum
    SeparateurLigne = pSeparateurLigne
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : type d'accès au fichier                                                               *
' *---------------------------------------------------------------------------------------------------*
Property Let TypeAcces(TypeAcces As ConnectModeEnum)

    If TypeAcces = AD_MODE_WRITE Or TypeAcces = AD_MODE_READ Then
        pTypeAcces = TypeAcces
    Else
        Err.Raise -1040, , "Le type d'accès " & TypeAcces & " est invalide."
    End If
    
End Property

Property Get TypeAcces() As ConnectModeEnum
    TypeAcces = pTypeAcces
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : nom du fichier                                                                        *
' *---------------------------------------------------------------------------------------------------*
Property Let NomFichier(NomFichier As String)

    If Trim(NomFichier) = "" Then
        Err.Raise -1050, , "Le nom du fichier doit être renseigné."
    Else
        ' Convertir les noms de fichier sous forme d'URL au format Lecteur:\Répertoire\NomFichier.extension
        pNomFichier = ConvertirUrlSharePoint(NomFichier)
    End If
    
End Property

Property Get NomFichier() As String
    NomFichier = pNomFichier
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : titre de la boîte de dialogue à afficher                                              *
' *---------------------------------------------------------------------------------------------------*
Property Let TitreBoiteDeDialogue(Titre As String)

    If Trim(Titre) = "" Then
        Err.Raise -1060, , "Le titre de la boîte de dialogue doit être renseigné."
    Else
        pTitreBoiteDialogue = Titre
    End If
    
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : libellé du filtre à appliquer afin de sélectionner les fichiers                       *
' *---------------------------------------------------------------------------------------------------*
Property Let LibelleFiltre(sLibelleFiltre As String)

    If Trim(sLibelleFiltre) = "" Then
        Err.Raise -1070, , "Le libellé du filtre doit être renseigné."
    Else
        pLibelleFiltre = sLibelleFiltre
    End If
    
End Property

Property Get LibelleFiltre() As String
    LibelleFiltre = pLibelleFiltre
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : Filtre des extensions séparées par une virgule, eg. *.csv,*.txt                       *
' *---------------------------------------------------------------------------------------------------*
Property Let ExtensionFiltre(sExtensionFiltre As String)

    If Trim(sExtensionFiltre) = "" Then
        Err.Raise -1080, , "Les extensions doivent être renseignées."
    Else
        pExtensionFiltre = sExtensionFiltre
    End If
    
End Property

Property Get ExtensionFiltre() As String
    ExtensionFiltre = pExtensionFiltre
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : nom du fichier pré renseigné dans la boite de dialogue pour enregistrer sous          *
' *---------------------------------------------------------------------------------------------------*
Property Let NomInitialFichier(NomInitialFichier As String)

    If Trim(NomInitialFichier) = "" Then
        Err.Raise -1090, , "Le nom initial du fichier n'est pas renseigné."
    Else
        pNomInitialFichier = NomInitialFichier
    End If

End Property

Property Get NomInitialFichier() As String
    NomInitialFichier = pNomInitialFichier
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : chemin d'un répertoire (dossier)                                                      *
' *---------------------------------------------------------------------------------------------------*
Property Let Repertoire(Repertoire As String)
    
    If Trim(Repertoire) = "" Then
        Err.Raise -1100, , "Le répertoire n'est pas renseigné."
    End If
    
    pRepertoire = Repertoire
    
End Property

Property Get Repertoire() As String
    Repertoire = pRepertoire
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : Filtre à appliquer lors de la sélection des fichiers à enregistrer sous               *
' *---------------------------------------------------------------------------------------------------*
Property Let Filtre(lIndexFiltre As FilterIndexEnum)

    If lIndexFiltre < 0 Or lIndexFiltre > 29 Then
        Err.Raise -1110, , "Le filtre " & lIndexFiltre & " est invalide."
    Else
        pIndexFiltre = lIndexFiltre
    End If
    
End Property

Property Get Filtre() As FilterIndexEnum
    Filtre = pIndexFiltre
End Property

' *---------------------------------------------------------------------------------------------------*
' * Ouverture d'un fichier via ADODB                                                                  *
' *---------------------------------------------------------------------------------------------------*
' * Propriétés utilisées :                                                                            *
' *   -- TypeFichier     (obligatoire)                                                                *
' *   -- SeparateurLigne (obligatoire si TypeFichier = Texte et lecture par ligne)                    *
' *   -- Encodage        (obligatoire si TypeFichier = Texte)                                         *
' *   -- TypeAcces       (obligatoire)                                                                *
' *   -- NomFichier      (obligatoire si accès en lecture)                                            *
' *---------------------------------------------------------------------------------------------------*
Public Sub Ouvrir()
    
    Select Case True
        ' Vérifier le type de fichier
        Case pTypeFichier = NON_DEFINI
            Err.Raise -1120, , "Le type de fichier n'est pas défini."
        ' Vérifier que l'encodage est renseigné si le fichier est de type Texte
        Case pTypeFichier = AD_TYPE_TEXT And pEncodage = ""
            Err.Raise -1121, , "L'encodage n'est pas défini."
        ' Vérifier que le type d'accès est renseigné
        Case pTypeAcces = NON_DEFINI
            Err.Raise -1122, , "Le type d'accès n'est pas défini."
        Case pTypeAcces = AD_MODE_READ
            ' Si le nom de fichier n'est pas renseigné
            If Trim(pNomFichier) = "" Then
                Err.Raise -1123, , "Le nom du fichier n'est pas défini."
            End If
            ' Vérifier que le fichier existe si l'accès est en lecture
            If FichierExiste = False Then
                Err.Raise -1124, , "Le fichier " & pNomFichier & " n'existe pas."
            End If
    End Select
    
    With pFichierADODB
        ' Type de fichier : binaire ou texte
        .Type = pTypeFichier
        ' Si fichier texte, préciser l'encodage et le séparateur
        If pTypeFichier = AD_TYPE_TEXT Then
            .Charset = pEncodage
            If pSeparateurLigne <> NON_DEFINI Then
                .LineSeparator = pSeparateurLigne
            End If
        End If
        .Open
        ' Si ouverture en lecture alors on charge le fichier
        If pTypeAcces = AD_MODE_READ Then
            .LoadFromFile (pNomFichier)
        End If
    End With
    ' Fichier ouvert sans erreur
    pStatutFichier = FICHIER_OUVERT
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Lecture des n prochains caractères d'un fichier texte ADODB                                       *
' *---------------------------------------------------------------------------------------------------*
' * Paramètre utilisé :                                                                               *
' *   -- lNbreCarALire    : nombre de caractères à lire dans le fichier ou jusqu'à la fin de fichier  *
' *---------------------------------------------------------------------------------------------------*
Public Function Lire(lNbreCarALire As Long) As String
    
    Select Case True
        ' Accès impossible à un fichier non ouvert
        Case pStatutFichier <> FICHIER_OUVERT
            Err.Raise -1130, , "Le fichier n'a pas été ouvert."
        ' Le nombre de caractères à lire doit être > 0
        Case lNbreCarALire <= 0
            Err.Raise -1131, , "Le nombre de caractères à lire n'est pas valide."
        ' Seule la lecture complète d'un fichier binaire est possible
        Case pTypeFichier = AD_TYPE_BINARY
            Err.Raise -1132, , "La lecture partielle d'un fichier binaire n'est pas possible."
        Case Else
            ' Lire les n caractères suivants
            Lire = pFichierADODB.ReadText(lNbreCarALire)
            pNbreOctectsLus = pNbreOctectsLus + Len(Lire)
    End Select
    
End Function

' *---------------------------------------------------------------------------------------------------*
' * Lecture du prochain enregistrement d'un fichier texte ADODB                                       *
' *---------------------------------------------------------------------------------------------------*
' * Propriété utilisée :                                                                              *
' *   -- SeparateurLigne                                                                              *
' *---------------------------------------------------------------------------------------------------*
Public Function LireEnregistrement() As String
    
    Select Case True
        ' Accès impossible à un fichier non ouvert
        Case pStatutFichier <> FICHIER_OUVERT
            Err.Raise -1140, , "Le fichier n'a pas été ouvert."
        ' Vérifie que le séparateur a bien été défini pour la lecture par ligne
        Case pSeparateurLigne = NON_DEFINI
            Err.Raise -1141, , "Le séparateur de ligne n'est pas défini."
        ' Lecture du prochain enregistrement
        Case pTypeFichier = AD_TYPE_TEXT
            LireEnregistrement = pFichierADODB.ReadText(AD_READ_LINE)
            pNbreEnregLus = pNbreEnregLus + 1
            pNbreOctectsLus = pNbreOctectsLus + Len(LireEnregistrement)
        ' Les fichiers binaires ne comportent pas de séparateur, par conséquent, il n'est pas possible de lire un enregistrement
        Case pTypeFichier = AD_TYPE_BINARY
            Err.Raise -1142, , "La lecture d'un enregistrement d'un fichier binaire n'est pas possible."
    End Select
    
End Function

' *---------------------------------------------------------------------------------------------------*
' * Fin de fichier atteinte ?                                                                         *
' *---------------------------------------------------------------------------------------------------*
Public Function FinFichier() As Boolean
    
    ' Accès impossible à un fichier non ouvert
    If pStatutFichier <> FICHIER_OUVERT Then
        Err.Raise -1150, , "Le fichier n'a pas été ouvert."
    End If
    
    FinFichier = pFichierADODB.EOS

End Function

' *---------------------------------------------------------------------------------------------------*
' * Position dans le fichier                                                                          *
' *---------------------------------------------------------------------------------------------------*
Public Sub Position(Position As Long)

    pFichierADODB.Position = Position
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Lecture intégrale d'un fichier ADODB                                                              *
' *---------------------------------------------------------------------------------------------------*
Public Function LireFichier() As Variant
    
    Select Case True
        ' Accès impossible à un fichier non ouvert
        Case pStatutFichier <> FICHIER_OUVERT
            Err.Raise -1160, , "Le fichier n'a pas été ouvert."
        ' Lire entièrement un fichier binaire
        Case pTypeFichier = AD_TYPE_BINARY
            LireFichier = pFichierADODB.Read(AD_READ_ALL)
            pNbreOctectsLus = pNbreOctectsLus + Len(LireFichier)
        ' Lire entièrement un fichier texte
        Case pTypeFichier = AD_TYPE_TEXT
            LireFichier = pFichierADODB.ReadText(AD_READ_ALL)
            pNbreOctectsLus = pNbreOctectsLus + Len(LireFichier)
    End Select

End Function

' *---------------------------------------------------------------------------------------------------*
' * Ecrire un enregistrement dans un flux ADODB.stream                                                *
' *---------------------------------------------------------------------------------------------------*
' * Paramètres :                                                                                      *
' *   -- Enregistrement (Obligatoire) : enregistrement à écrire (XML, json, texte, etc)               *
' *---------------------------------------------------------------------------------------------------*
Public Sub EcrireEnregistrement(Enregistrement As String)
    
    Select Case True
        ' Accès impossible à un fichier non ouvert
        Case pStatutFichier <> FICHIER_OUVERT
            Err.Raise -1170, , "Le fichier n'a pas été ouvert."
        ' Ecriture de l'enregistrement dans le fichier texte. Les séparateurs CR/LF doivent être présents dans l'enregistrement.
        Case pTypeFichier = AD_TYPE_TEXT
            pFichierADODB.WriteText Enregistrement
            pNbreEnregEcrits = pNbreEnregEcrits + 1
            pNbreOctectsEcrits = pNbreOctectsEcrits + Len(Enregistrement)
        ' Il n'est pas possible d'écrire un enregistrement dans un fichier binaire.
        Case pTypeFichier = AD_TYPE_BINARY
            Err.Raise -1171, , "L'écriture d'un enregistrement dans un fichier binaire n'est pas possible."
    End Select
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Ecrire des données dans un flux ADODB.stream                                                      *
' *---------------------------------------------------------------------------------------------------*
' * Paramètres :                                                                                      *
' *   -- Donnees (Obligatoire) : données à écrire (image, pdf, etc)                                   *
' *---------------------------------------------------------------------------------------------------*
Public Sub Ecrire(Donnees As Variant)
    
    Select Case True
        ' Accès impossible à un fichier non ouvert
        Case pStatutFichier <> FICHIER_OUVERT
            Err.Raise -1180, , "Le fichier n'a pas été ouvert."
        ' Ecrire le fichier binaire (en une seule fois)
        Case pTypeFichier = AD_TYPE_BINARY
            pFichierADODB.Write Donnees
            pNbreOctectsEcrits = pNbreOctectsEcrits + Len(Donnees)
        ' Il n'est pas possible d'écrire des données brutes dans un fichier texte.
        Case pTypeFichier = AD_TYPE_TEXT
            Err.Raise -1181, , "L'écriture de données binaires dans fichier texte n'est pas possible."
    End Select
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Enregistrer le flux ADODB.Stream sous le nom du fichier                                           *
' *---------------------------------------------------------------------------------------------------*
' * Propriétés utilisées :                                                                            *
' *   -- TypeFichier     (obligatoire)                                                                *
' *   -- NomFichier      (obligatoire)                                                                *
' *---------------------------------------------------------------------------------------------------*
Public Sub EnregistrerSous()

    Select Case True
        Case pStatutFichier <> FICHIER_OUVERT
            ' Accès impossible à un fichier non ouvert
            Err.Raise -1190, , "Le fichier n'a pas été ouvert."
        Case Trim(pNomFichier) = ""
            Err.Raise -1191, , "Le nom du fichier n'est pas défini."
        Case pTypeFichier = NON_DEFINI
            Err.Raise -1192, , "Le type de fichier n'est pas défini."
        Case Else
            If pEncodageNum = AD_UTF_8 Then
                ' Création d'un flux pour le fichier UTF_8 sans BOM
                Dim oFichierUtf8 As Object
                Set oFichierUtf8 = CreateObject("ADODB.STREAM")
                With oFichierUtf8
                    .Mode = AD_MODE_READ_WRITE
                    .Type = AD_TYPE_BINARY
                    .Open
                End With
                With pFichierADODB
                    ' Ignorer les caractères du BOM
                    .Position = 3
                    ' Copier le contenu du flux dans le nouveau flux
                    .CopyTO oFichierUtf8
                    .Flush
                End With
                With oFichierUtf8
                    .SaveToFile pNomFichier, AD_SAVE_CREATE_OVER_WRITE
                    .Close
                End With
                Set oFichierUtf8 = Nothing
            Else
                pFichierADODB.SaveToFile pNomFichier, AD_SAVE_CREATE_OVER_WRITE
            End If
    End Select
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Fermer un objet ADODB                                                                             *
' *---------------------------------------------------------------------------------------------------*
Public Sub Fermer()
    
    If pStatutFichier <> FICHIER_OUVERT Then
        ' Accès impossible à un fichier non ouvert
        Err.Raise -1200, , "Le fichier n'a pas été ouvert."
    Else
        pFichierADODB.Close
        pStatutFichier = FICHIER_FERME
    End If
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Compteurs                                                                                         *
' *---------------------------------------------------------------------------------------------------*

Public Function NbreEnregLus() As Long
    NbreEnregLus = pNbreEnregLus
End Function

Public Function NbreEnregEcrits() As Long
    NbreEnregEcrits = pNbreEnregEcrits
End Function

Public Function NbreOctectsLus() As Long
    NbreOctectsLus = pNbreOctectsLus
End Function

Public Function NbreOctectsEcrits() As Long
    NbreOctectsEcrits = pNbreOctectsEcrits
End Function

' *---------------------------------------------------------------------------------------------------*
' * Affiche la boite de dialogue qui permet de sélectionner le fichier afin d'enregistrer sous        *
' *---------------------------------------------------------------------------------------------------*
' * Propriétés utilisées :                                                                            *
' *   -- Filtre              (Facultatif)                                                             *
' *   -- TitreBoiteDialogue  (Facultatif)                                                             *
' *---------------------------------------------------------------------------------------------------*
Public Sub SelectionnerFichierEnregistrerSous()

    With Application.FileDialog(msoFileDialogSaveAs)
        .AllowMultiSelect = False
        If pNomInitialFichier <> "" Then
            .InitialFileName = pNomInitialFichier
        End If
        .InitialView = msoFileDialogViewDetails
        If pIndexFiltre <> NON_DEFINI Then
            .FilterIndex = pIndexFiltre
        End If
        ' Titre par défaut de la boîte de dialogue
        If pTitreBoiteDialogue = "" Then
            .Title = "Sélectionner le répertoire et saisir le nom du fichier à enregistrer sous"
        Else
            .Title = pTitreBoiteDialogue
        End If
        If .Show = -1 Then
            pNomFichier = .SelectedItems(1)
        Else
            pNomFichier = ""
        End If
    End With

End Sub

' *---------------------------------------------------------------------------------------------------*
' * Affiche la boite de dialogue qui permet de sélectionner un fichier qui correspond aux filtres     *
' *---------------------------------------------------------------------------------------------------*
' * Propriétés utilisées :                                                                            *
' *   -- LibelleFiltre       (Facultatif)                                                             *
' *   -- ExtensionFiltre     (Facultatif)                                                             *
' *   -- TitreBoiteDialogue  (Facultatif)                                                             *
' *---------------------------------------------------------------------------------------------------*
Public Sub SelectionnerFichier()

    ' Boite de dialogue de sélection d'un fichier
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        If pNomInitialFichier <> "" Then
            .InitialFileName = pNomInitialFichier
        End If
        .InitialView = msoFileDialogViewDetails
        ' Si filtre ou extension non renseigné
        If pIndexFiltre <> NON_DEFINI Then
            .FilterIndex = pIndexFiltre
        Else
            If pLibelleFiltre = "" Or pExtensionFiltre = "" Then
                ' Tous les fichiers par défaut
                .FilterIndex = 1
            Else
                .Filters.Add pLibelleFiltre, pExtensionFiltre, 1
                .FilterIndex = 0
            End If
        End If
        ' Titre par défaut si non renseigné
        If pTitreBoiteDialogue = "" Then
            .Title = "Sélectionner le répertoire et le fichier"
        Else
            .Title = pTitreBoiteDialogue
        End If
        If .Show = -1 Then
            pNomFichier = .SelectedItems(1)
        Else
            pNomFichier = ""
        End If
    End With
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' Affiche la boite de dialogue qui permet de sélectionner un répertoire                               *
' *---------------------------------------------------------------------------------------------------*
' * Propriété utilisée :                                                                              *
' *   -- TitreBoiteDialogue  (Facultatif)                                                             *
' *---------------------------------------------------------------------------------------------------*
Public Sub SelectionnerRepertoire()

    ' Boite de dialogue de sélection du répertoire
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        If pNomInitialFichier <> "" Then
            .InitialFileName = pNomInitialFichier
        End If
        .InitialView = msoFileDialogViewDetails
        If pTitreBoiteDialogue = "" Then
            .Title = "Sélectionner le répertoire"
        Else
            .Title = pTitreBoiteDialogue
        End If
        ' Si aucun répertoire sélectionné alors fin du traitement
        If .Show = -1 Then
            pRepertoire = .SelectedItems(1)
        Else
            pRepertoire = ""
        End If
    End With

End Sub

' *---------------------------------------------------------------------------------------------------*
' * Vérifie si un fichier est trop volumineux pour être traité par ADODB                              *
' *---------------------------------------------------------------------------------------------------*
' * Propriété utilisée :                                                                              *
' *   -- NomFichier          (Obligatoire)                                                            *
' *---------------------------------------------------------------------------------------------------*
Public Function EstFichierVolumineux(Optional NomFichier As String) As Boolean
    
    Dim NomFichierLocal As String
    
    Select Case True
        Case Trim(NomFichier) <> ""
            NomFichierLocal = NomFichier
        Case pNomFichier <> ""
            NomFichierLocal = pNomFichier
        Case Else
            Err.Raise -1210, , "Le nom du fichier n'est pas défini."
    End Select
    
    If LongueurFichier(NomFichierLocal) > 4400000000# Then
        EstFichierVolumineux = True
    Else
        EstFichierVolumineux = False
    End If

End Function

' *---------------------------------------------------------------------------------------------------*
' * Retourne la longueur d'un fichier (en octets)                                                     *
' *---------------------------------------------------------------------------------------------------*
' * Paramètre utilisé :                                                                               *
' *   -- NomFichier         (Facultatif)                                                              *
' * Propriété utilisée :                                                                              *
' *   -- NomFichier          (Facultatif)                                                             *
' *---------------------------------------------------------------------------------------------------*
Public Function LongueurFichier(Optional NomFichier As String) As Double

    Dim NomFichierLocal As String
    Dim oFileSystem As Object, oFichier As Object
    
    Select Case True
        Case Trim(NomFichier) <> ""
            NomFichierLocal = NomFichier
        Case pNomFichier <> ""
            NomFichierLocal = pNomFichier
        Case Else
            Err.Raise -1230, , "Le nom du fichier n'est pas défini."
    End Select

    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
    Set oFichier = oFileSystem.GetFile(NomFichierLocal)
    LongueurFichier = oFichier.Size

End Function

' *---------------------------------------------------------------------------------------------------*
' * Vérifie si un fichier existe                                                                      *
' *---------------------------------------------------------------------------------------------------*
' * Paramètre utilisé :                                                                               *
' *   -- NomFichier         (Facultatif)                                                              *
' * Propriété utilisée :                                                                              *
' *   -- NomFichier          (Facultatif)                                                             *
' *---------------------------------------------------------------------------------------------------*
Public Function FichierExiste(Optional NomFichier As String) As Boolean
    
    Dim NomFichierLocal As String
    
    Select Case True
        Case Trim$(NomFichier) <> ""
            NomFichierLocal = NomFichier
        Case pNomFichier <> ""
            NomFichierLocal = pNomFichier
        Case Else
            Err.Raise -1240, , "Le nom du fichier n'est pas défini."
    End Select

    FichierExiste = Dir(NomFichierLocal, vbNormal) <> ""

End Function

' *---------------------------------------------------------------------------------------------------*
' * Vérifie si un répertoire existe                                                                   *
' *---------------------------------------------------------------------------------------------------*
' * Propriété utilisée :                                                                              *
' *   -- Repertoire          (Obligatoire)                                                            *
' *---------------------------------------------------------------------------------------------------*
Public Function RepertoireExiste(Optional Repertoire As String) As Boolean
    
    Dim NomRepertoire As String
    
    Select Case True
        Case Trim(Repertoire) <> ""
            NomRepertoire = Repertoire
        Case pRepertoire <> ""
            NomRepertoire = pRepertoire
        Case Else
            Err.Raise -1250, , "Le nom du répertoire n'est pas défini."
    End Select

    RepertoireExiste = Dir(NomRepertoire, vbDirectory) <> ""

End Function

' *--------------------------------------------------------------------------------------------------------------------------*
' * Convertir un nom de chemin défini par une URL OneDrive ou SharePoint vers un nom de chemin Windows                       *
' * Exemple : https://xxx-my.sharepoint.com/personal/ devient c:\Users\xxxx\OneDrive - xxx                                   *
' *--------------------------------------------------------------------------------------------------------------------------*
Private Function ConvertirUrlSharePoint(Chemin As String) As String

    Dim sListeDossiers() As String, iNbDossiers As Integer, lPosDoc As Long, Repertoire As String
    
    ' Si le chemin du fichier commence par http
    If LCase$(Left(Chemin, 4)) = "http" Then
        Select Case True
        ' Espace personnel sur SharePoint (i.e. OneDrive Commercial) ?
        Case Chemin Like "https://*-my.sharepoint.com/personal/*"
            ' Recherche la chaîne "/Documents/documents" afin d'obtenir le début de l'arborescence dans le dossier des documents
            lPosDoc = InStr(1, Chemin, "/Documents/Documents/", vbTextCompare) + Len("/Documents")
            ' Le répertoire local est récupéré à partir du 2ème /Documents
            Repertoire = Mid(Chemin, lPosDoc, Len(Chemin) - lPosDoc + 1)
            ConvertirUrlSharePoint = Environ("OneDriveCommercial") & Replace(Repertoire, "/", "\")
        ' Espace de travail partagé
        Case Chemin Like "https://weshare*"
            sListeDossiers = Split(Chemin, "/")
            ConvertirUrlSharePoint = "\\" & sListeDossiers(2) & "@SSL\DavWWWRoot"
            For iNbDossiers = 3 To UBound(sListeDossiers)
                ConvertirUrlSharePoint = ConvertirUrlSharePoint & "\" & sListeDossiers(iNbDossiers)
            Next
        Case Chemin Like "https://d.docs.live.net/*"
            ' Recherche la chaîne "/documents" afin d'obtenir le début de l'arborescence dans le dossier des documents
            lPosDoc = InStr(1, Chemin, "/Documents/", vbTextCompare)
            ' Le répertoire local est récupéré à partir du 2ème /Documents
            Repertoire = Mid(Chemin, lPosDoc, Len(Chemin) - lPosDoc + 1)
            ConvertirUrlSharePoint = Environ("OneDrive") & Replace(Repertoire, "/", "\")
        End Select
    Else
        ConvertirUrlSharePoint = Chemin
    End If
    
End Function

' *---------------------------------------------------------------------------------------------------*
' * Retourne l'extension du fichier                                                                   *
' *---------------------------------------------------------------------------------------------------*
Public Function Extension() As String
    
    Dim lPosPt As Long
    
    lPosPt = InStrRev(pNomFichier, ".")
    If lPosPt > 0 Then
        Extension = LCase$(Right$(pNomFichier, Len(pNomFichier) - lPosPt))
    Else
        Extension = ""
    End If
End Function

' *---------------------------------------------------------------------------------------------------*
' * Initialisation de la classe : création des objets ADODB.Stream et le dictionnaire des encodages   *
' *---------------------------------------------------------------------------------------------------*
Private Sub Class_Initialize()
    
    ' Création d'un objet ADODB
    Set pFichierADODB = CreateObject("ADODB.STREAM")

    ' Création d'un dictionnaire des encodages
    Set DicEncodage = New Dictionary
    ' Ajout des encodages (à compléter si nécessaire)
    ' clé : libellé de l'encodage identique aux clés de registre charset de windows
    ' La liste ci-dessous doit être identique à l'enum CharsetEnum
    DicEncodage.Add NON_DEFINI, "Non défini"
    DicEncodage.Add AD_ASCII, "iso-8859-1"
    DicEncodage.Add AD_LATIN_2, "iso-8859-2"
    DicEncodage.Add AD_LATIN_4, "iso-8859-4"
    DicEncodage.Add AD_CYRILLIQUE, "iso-8859-5"
    DicEncodage.Add AD_GREC, "iso-8859-7"
    DicEncodage.Add AD_LATIN_5, "iso-8859-9"
    DicEncodage.Add AD_COREEN, "ks_c_5601-1987"
    DicEncodage.Add AD_UTF_7, "utf-7"
    DicEncodage.Add AD_UTF_8, "utf-8"
    DicEncodage.Add AD_UTF_8_BOM, "utf-8"
    DicEncodage.Add AD_CP1250, "Windows-1250"
    DicEncodage.Add AD_CP1251, "Windows-1251"
    DicEncodage.Add AD_X_ANSI, "windows-1252"
    DicEncodage.Add AD_ISO_8859_11, "windows-874"
    DicEncodage.Add AD_ARABE, "iso-8859-6"
    DicEncodage.Add AD_AUTRE, "Autre"
    
    ' Fichier pas encore ouvert
    pStatutFichier = NON_DEFINI
    pNbreEnregLus = 0
    pNbreEnregEcrits = 0
    pNbreOctectsLus = 0
    pNbreOctectsEcrits = 0
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Terminaison de la classe : libérer la mémoire en détruisant les objets                            *
' *---------------------------------------------------------------------------------------------------*
Private Sub Class_Terminate()
    Set pFichierADODB = Nothing
    Set DicEncodage = Nothing
End Sub
