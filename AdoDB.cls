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
' Copyright (c) 2026, Vincent ROSSET
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

'-------------------------------------------------------------------------------------------------------------------------
' AdoDB : Classe VBA de gestion des flux ADODB (uniquement vers des fichiers)
'-------------------------------------------------------------------------------------------------------------------------

' Nécessite d'activer la référence "Microsoft Scripting RunTime"
'
' Dans l'éditeur de macros (Alt+F11): Menu Outils \ Références
' Cochez les lignes :
' -- "Microsoft Scripting RunTime"
' Cliquez sur le bouton OK pour valider.

Option Explicit
Option Compare Text

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'
' Propriétés de la classe :
' -------------------------
'   TypeFichier                        : Type du fichier (texte ou binaire)
'   Encodage                           : Encodage du fichier (enum)
'   EncodageTxt                        : Encodage du fichier (saisie libre)
'   SeparateurLigne                    : Séparateur de lignes d'un fichier texte
'   ModeAcces                          : Accès au fichier (lecture, écriture)
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

Const NON_DEFINI = 0

'-------------------------------------------------------------------------------------------------------------------------
' Enum du statut du flux AdoDB
'-------------------------------------------------------------------------------------------------------------------------
Private Enum FileStatusEnum
    UNDEFINED_FILE_STATUS = 0
    OPENED_FILE = 1
    CLOSED_FILE = 2
End Enum

'-------------------------------------------------------------------------------------------------------------------------
' Propriétés des fichiers
'-------------------------------------------------------------------------------------------------------------------------

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
Private pModeAcces As ConnectModeEnum

'-------------------------------------------------------------------------------------------------------------------------
' Gestion des boîtes de dialogue
'-------------------------------------------------------------------------------------------------------------------------

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

'-------------------------------------------------------------------------------------------------------------------------
' Gestion des fichiers
'-------------------------------------------------------------------------------------------------------------------------

' Statut du fichier
Private pStatutFichier As FileStatusEnum
' Compteurs d'enregistrements
Private pNbreEnregLus As Long
Private pNbreEnregEcrits As Long
' Compteurs en octets
Private pNbreOctectsLus As Long
Private pNbreOctectsEcrits As Long

'-------------------------------------------------------------------------------------------------------------------------
' Pointer sur un objet AdoDB externe à la classe
'-------------------------------------------------------------------------------------------------------------------------
' FichierAdoDB : Objet AdoDB sur lequel pointer
'-------------------------------------------------------------------------------------------------------------------------
Property Set Fichier(FichierAdodb As Object)

    Set pFichierADODB = FichierAdodb

End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : Type de fichier
'-------------------------------------------------------------------------------------------------------------------------
' TypeFichier : Type de fichier (texte ou binaire)
'-------------------------------------------------------------------------------------------------------------------------
Property Let TypeFichier(TypeFichier As StreamTypeEnum)

    Select Case TypeFichier
        Case AD_TYPE_BINARY
            Select Case True
                Case pEncodage <> ""
                    ' L'encodage n'est pas nécessaire pour lire un fichier binaire
                    Err.Raise -100, "AdoDB.TypeFichier", "L'encodage ne doit pas être défini pour les fichiers binaires."
                Case pSeparateurLigne <> NON_DEFINI
                    ' Le séparateur de lignes n'est pas nécessaire pour lire un fichier binaire
                    Err.Raise -101, "AdoDB.TypeFichier", "Le séparateur de lignes ne doit pas être défini pour les fichiers binaires."
                Case Else
                    pTypeFichier = TypeFichier
            End Select
        Case AD_TYPE_TEXT
            pTypeFichier = TypeFichier
        Case Else
            Err.Raise -102, "AdoDB.TypeFichier", "Le type de fichier " & TypeFichier & " est invalide."
    End Select

End Property

Property Get TypeFichier() As StreamTypeEnum
    TypeFichier = pTypeFichier
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : Encodage (Enum)
'-------------------------------------------------------------------------------------------------------------------------
' EncodageNum : valeur numérique définie dans CharsetEnum de l'encodage du fichier
'-------------------------------------------------------------------------------------------------------------------------
Property Let Encodage(EncodageNum As CharsetEnum)

    Select Case True
        ' L'encodage doit être renseigné
        Case EncodageNum = NON_DEFINI
            Err.Raise -110, "AdoDB.Encodage", "L'encodage n'est pas renseigné."
        Case pTypeFichier = AD_TYPE_BINARY
            ' L'encodage n'est pas nécessaire pour lire un fichier binaire
            Err.Raise -111, "AdoDB.Encodage", "L'encodage ne doit pas être défini pour les fichiers binaires."
            ' Vérifie que l'encodage existe dans l'enum
        Case DicEncodage.Exists(EncodageNum)
            pEncodageNum = EncodageNum
            pEncodage = DicEncodage(EncodageNum)
        Case Else
            ' Si l'encodage est absent de l'enum ListeEncodage et du dictionnaire dicEncodage alors une erreur est retournée.
            ' La liste des encodages n'est pas exhaustive et doit être complétée si nécessaire, sinon utiliser la propriété EncodageTxt.
            Err.Raise -112, "AdoDB.Encodage", "L'encodage n'existe pas dans l'enum CharsetEnum."
    End Select
    
End Property

Property Get Encodage() As CharsetEnum
    
    Encodage = pEncodageNum
    
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : Encodage au format texte
'-------------------------------------------------------------------------------------------------------------------------
' Encodage : libellé de l'encodage à utiliser pour accéder au fichier (utf-8, ascii, etc)
'-------------------------------------------------------------------------------------------------------------------------
' Aucun contrôle n'est réalisé, l'encodage est transmis tel quel à AdoDB qui va retourner une erreur s'il n'est pas valide.
'-------------------------------------------------------------------------------------------------------------------------
Property Let EncodageTxt(Encodage As String)

    Select Case True
        ' L'encodage doit être renseigné
        Case Trim(Encodage) = ""
            Err.Raise -120, "AdoDB.EncodageTxt", "L'encodage n'est pas renseigné."
        ' L'encodage n'est pas nécessaire pour lire un fichier binaire
        Case pTypeFichier = AD_TYPE_BINARY
            Err.Raise -121, "AdoDB.EncodageTxt", "L'encodage ne doit pas être défini pour les fichiers binaires."
        ' Aucun contrôle de l'encodage, la saisie est libre.
        Case Else
            pEncodage = Encodage
            pEncodageNum = AD_AUTRE
    End Select
    
End Property

Property Get EncodageTxt() As String
    EncodageTxt = pEncodage
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : Séparateur de lignes (caractères de fin de ligne)
'-------------------------------------------------------------------------------------------------------------------------
' SeparateurLigne : valeur numérique définie dans LineSeparatorEnum du séparateur de ligne
'-------------------------------------------------------------------------------------------------------------------------
Property Let SeparateurLigne(SeparateurLigne As LineSeparatorsEnum)

    ' Vérifier la valeur du séparateur de ligne
    If SeparateurLigne = AD_CR Or SeparateurLigne = AD_CR_LF Or SeparateurLigne = AD_LF Then
        If pTypeFichier = AD_TYPE_BINARY Then
            ' Le séparateur ne sert à rien pour les fichiers binaires
            Err.Raise -130, "AdoDB.SeparateurLigne", "Le séparateur ne doit pas être défini pour les fichiers binaires."
        Else
            pSeparateurLigne = SeparateurLigne
        End If
    Else
        Err.Raise -131, "AdoDB.SeparateurLigne", "Le séparateur " & SeparateurLigne & " est invalide."
    End If
    
End Property

Property Get SeparateurLigne() As LineSeparatorsEnum
    SeparateurLigne = pSeparateurLigne
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : mode d'accès au flux
'-------------------------------------------------------------------------------------------------------------------------
' ModeAcces : valeur numérique définie dans ConnectModeEnum du mode d'accès à un flux AdoDB
'-------------------------------------------------------------------------------------------------------------------------
Property Let ModeAcces(ModeAcces As ConnectModeEnum)

    If ModeAcces = AD_MODE_WRITE Or ModeAcces = AD_MODE_READ Or ModeAcces = AD_MODE_READ_WRITE Then
        pModeAcces = ModeAcces
    Else
        Err.Raise -140, "AdoDB.ModeAcces", "Le type d'accès " & ModeAcces & " est invalide."
    End If
    
End Property

Property Get ModeAcces() As ConnectModeEnum
    ModeAcces = pModeAcces
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : nom physique du fichier à traiter
'-------------------------------------------------------------------------------------------------------------------------
Property Let NomFichier(NomFichier As String)

    If Trim(NomFichier) = "" Then
        Err.Raise -150, "AdoDB.Nomfichier", "Le nom du fichier doit être renseigné."
    Else
        ' Convertir les noms de fichier sous forme d'URL au format Lecteur:\Répertoire\NomFichier.extension
        pNomFichier = ConvertirUrlSharePoint(NomFichier)
    End If
    
End Property

Property Get NomFichier() As String
    NomFichier = pNomFichier
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : titre de la boîte de dialogue qui permet de sélectionner un fichier / répertoire
'-------------------------------------------------------------------------------------------------------------------------
Property Let TitreBoiteDeDialogue(Titre As String)

    If Trim(Titre) = "" Then
        Err.Raise -160, "AdoDB.TitreBoiteDeDialogue", "Le titre de la boîte de dialogue doit être renseigné."
    Else
        pTitreBoiteDialogue = Titre
    End If
    
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : libellé du filtre à appliquer afin de sélectionner les fichiers
'-------------------------------------------------------------------------------------------------------------------------
Property Let LibelleFiltre(sLibelleFiltre As String)

    If Trim(sLibelleFiltre) = "" Then
        Err.Raise -170, "AdoDB.LibelleFiltre", "Le libellé du filtre doit être renseigné."
    Else
        pLibelleFiltre = sLibelleFiltre
    End If
    
End Property

Property Get LibelleFiltre() As String
    LibelleFiltre = pLibelleFiltre
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : Filtre des extensions séparées par une virgule, eg. *.csv,*.txt
'-------------------------------------------------------------------------------------------------------------------------
Property Let ExtensionFiltre(sExtensionFiltre As String)

    If Trim(sExtensionFiltre) = "" Then
        Err.Raise -180, "AdoDB.ExtensionFiltre", "Les extensions doivent être renseignées."
    Else
        pExtensionFiltre = sExtensionFiltre
    End If
    
End Property

Property Get ExtensionFiltre() As String
    ExtensionFiltre = pExtensionFiltre
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : nom du fichier prérenseigné dans la boite de dialogue pour enregistrer sous
'-------------------------------------------------------------------------------------------------------------------------
Property Let NomInitialFichier(NomInitialFichier As String)

    If Trim(NomInitialFichier) = "" Then
        Err.Raise -190, "AdoDB.NomInitialFichier", "Le nom initial du fichier n'est pas renseigné."
    Else
        pNomInitialFichier = NomInitialFichier
    End If

End Property

Property Get NomInitialFichier() As String
    NomInitialFichier = pNomInitialFichier
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : chemin d'un répertoire (dossier)
'-------------------------------------------------------------------------------------------------------------------------
Property Let Repertoire(Repertoire As String)
    
    If Trim(Repertoire) = "" Then
        Err.Raise -200, "AdoDB.Repertoire", "Le répertoire n'est pas renseigné."
    End If
    
    pRepertoire = Repertoire
    
End Property

Property Get Repertoire() As String
    Repertoire = pRepertoire
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : Filtre à appliquer lors de la sélection des fichiers à enregistrer sous
'-------------------------------------------------------------------------------------------------------------------------
Property Let Filtre(lIndexFiltre As FilterIndexEnum)

    If lIndexFiltre < 0 Or lIndexFiltre > 29 Then
        Err.Raise -210, "AdoDB.Filtre", "Le filtre " & lIndexFiltre & " est invalide."
    Else
        pIndexFiltre = lIndexFiltre
    End If
    
End Property

Property Get Filtre() As FilterIndexEnum
    Filtre = pIndexFiltre
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Ouverture d'un fichier via ADODB
'-------------------------------------------------------------------------------------------------------------------------
' Poopriétés obligatoires :
'   TypeFichier
'   SeparateurLigne si TypeFichier = Texte et lecture par ligne
'   Encodage si TypeFichier = Texte
'   ModeAcces
'   NomFichier si accès en lecture
'-------------------------------------------------------------------------------------------------------------------------
Public Sub Ouvrir()
    
    Select Case True
        ' Vérifier le type de fichier
        Case pTypeFichier = NON_DEFINI
            Err.Raise -220, "AdoDB.Ouvrir", "Le type de fichier n'est pas défini."
        ' Vérifier que l'encodage est renseigné si le fichier est de type Texte
        Case pTypeFichier = AD_TYPE_TEXT And pEncodage = ""
            Err.Raise -221, "AdoDB.Ouvrir", "L'encodage n'est pas défini."
        ' Vérifier que le type d'accès est renseigné
        Case pModeAcces = NON_DEFINI
            Err.Raise -222, "AdoDB.Ouvrir", "Le type d'accès n'est pas défini."
        Case pModeAcces = AD_MODE_READ
            ' Si le nom de fichier n'est pas renseigné
            If Trim(pNomFichier) = "" Then
                Err.Raise -223, "AdoDB.Ouvrir", "Le nom du fichier n'est pas défini."
            End If
            ' Vérifier que le fichier existe si l'accès est en lecture
            If FichierExiste = False Then
                Err.Raise -224, "AdoDB.Ouvrir", "Le fichier " & pNomFichier & " n'existe pas."
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
        If pModeAcces = AD_MODE_READ Then
            .LoadFromFile (pNomFichier)
        End If
    End With
    ' Fichier ouvert sans erreur
    pStatutFichier = OPENED_FILE
    
End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Lecture des n prochains caractères d'un fichier texte ADODB
'-------------------------------------------------------------------------------------------------------------------------
' Paramètre(s) :
'   lNbreCarALire    : nombre de caractères à lire dans le fichier ou jusqu'à la fin de fichier
'-------------------------------------------------------------------------------------------------------------------------
Public Function Lire(lNbreCarALire As Long) As String
    
    Select Case True
        ' Accès impossible à un fichier non ouvert
        Case pStatutFichier <> OPENED_FILE
            Err.Raise -230, "AdoDB.Lire", "Le fichier n'a pas été ouvert."
        ' Le nombre de caractères à lire doit être > 0
        Case lNbreCarALire <= 0
            Err.Raise -231, "AdoDB.Lire", "Le nombre de caractères à lire n'est pas valide."
        ' Seule la lecture complète d'un fichier binaire est possible
        Case pTypeFichier = AD_TYPE_BINARY
            Err.Raise -232, "AdoDB.Lire", "La lecture partielle d'un fichier binaire n'est pas possible."
        Case Else
            ' Lire les n caractères suivants
            Lire = pFichierADODB.ReadText(lNbreCarALire)
            pNbreOctectsLus = pNbreOctectsLus + Len(Lire)
    End Select
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Lecture du prochain enregistrement d'un fichier texte ADODB
'-------------------------------------------------------------------------------------------------------------------------
' Propriété utilisée :
'   SeparateurLigne
'-------------------------------------------------------------------------------------------------------------------------
Public Function LireEnregistrement() As String
    
    Select Case True
        ' Accès impossible à un fichier non ouvert
        Case pStatutFichier <> OPENED_FILE
            Err.Raise -240, "AdoDB.LireEnregistrement", "Le fichier n'a pas été ouvert."
        ' Vérifie que le séparateur a bien été défini pour la lecture par ligne
        Case pSeparateurLigne = NON_DEFINI
            Err.Raise -241, "AdoDB.LireEnregistrement", "Le séparateur de ligne n'est pas défini."
        ' Lecture du prochain enregistrement
        Case pTypeFichier = AD_TYPE_TEXT
            LireEnregistrement = pFichierADODB.ReadText(AD_READ_LINE)
            pNbreEnregLus = pNbreEnregLus + 1
            pNbreOctectsLus = pNbreOctectsLus + Len(LireEnregistrement)
        ' Les fichiers binaires ne comportent pas de séparateur, par conséquent, il n'est pas possible de lire un enregistrement
        Case pTypeFichier = AD_TYPE_BINARY
            Err.Raise -242, "AdoDB.LireEnregistrement", "La lecture d'un enregistrement d'un fichier binaire n'est pas possible."
    End Select
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
'  Fin de fichier atteinte ?
'-------------------------------------------------------------------------------------------------------------------------
Public Function FinFichier() As Boolean
    
    ' Accès impossible à un fichier non ouvert
    If pStatutFichier <> OPENED_FILE Then
        Err.Raise -250, "AdoDB.FinFichier", "Le fichier n'a pas été ouvert."
    End If
    
    FinFichier = pFichierADODB.EOS

End Function

'-------------------------------------------------------------------------------------------------------------------------
' Position dans le fichier
'-------------------------------------------------------------------------------------------------------------------------
Public Sub Position(Position As Long)

    pFichierADODB.Position = Position
    
End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Lecture intégrale d'un fichier ADODB
'-------------------------------------------------------------------------------------------------------------------------
Public Function LireFichier() As Variant
    
    Select Case True
        ' Accès impossible à un fichier non ouvert
        Case pStatutFichier <> OPENED_FILE
            Err.Raise -260, "AdoDB.LireFichier", "Le fichier n'a pas été ouvert."
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

'-------------------------------------------------------------------------------------------------------------------------
' Ecrire un enregistrement dans un flux ADODB.stream
'-------------------------------------------------------------------------------------------------------------------------
' Paramètres :
'   Enregistrement (Obligatoire) : enregistrement à écrire (XML, json, texte, etc)
'-------------------------------------------------------------------------------------------------------------------------
Public Sub EcrireEnregistrement(Enregistrement As String)
    
    Select Case True
        ' Accès impossible à un fichier non ouvert
        Case pStatutFichier <> OPENED_FILE
            Err.Raise -270, "AdoDB.EcrireEnregistrement", "Le fichier n'a pas été ouvert."
        ' Ecriture de l'enregistrement dans le fichier texte. Les séparateurs CR/LF doivent être présents dans l'enregistrement.
        Case pTypeFichier = AD_TYPE_TEXT
            pFichierADODB.WriteText Enregistrement
            pNbreEnregEcrits = pNbreEnregEcrits + 1
            pNbreOctectsEcrits = pNbreOctectsEcrits + Len(Enregistrement)
        ' Il n'est pas possible d'écrire un enregistrement dans un fichier binaire.
        Case pTypeFichier = AD_TYPE_BINARY
            Err.Raise -271, "AdoDB.EcrireEnregistrement", "L'écriture d'un enregistrement dans un fichier binaire n'est pas possible."
    End Select
    
End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Ecrire des données dans un flux ADODB.stream
'-------------------------------------------------------------------------------------------------------------------------
' Paramètres :
'   Donnees (Obligatoire) : données à écrire (image, pdf, etc)
'-------------------------------------------------------------------------------------------------------------------------
Public Sub Ecrire(Donnees As Variant)
    
    Select Case True
        ' Accès impossible à un fichier non ouvert
        Case pStatutFichier <> OPENED_FILE
            Err.Raise -280, "AdoDB.Ecrire", "Le fichier n'a pas été ouvert."
        ' Ecrire le fichier binaire (en une seule fois)
        Case pTypeFichier = AD_TYPE_BINARY
            pFichierADODB.Write Donnees
            pNbreOctectsEcrits = pNbreOctectsEcrits + Len(Donnees)
        ' Il n'est pas possible d'écrire des données brutes dans un fichier texte.
        Case pTypeFichier = AD_TYPE_TEXT
            Err.Raise -281, "AdoDB.Ecrire", "L'écriture de données binaires dans fichier texte n'est pas possible."
    End Select
    
End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Enregistrer le flux ADODB.Stream sous le nom du fichier
'-------------------------------------------------------------------------------------------------------------------------
' Propriétés utilisées :
'   TypeFichier     (obligatoire)
'   NomFichier      (obligatoire)
'-------------------------------------------------------------------------------------------------------------------------
Public Sub EnregistrerSous()

    Select Case True
        Case pStatutFichier <> OPENED_FILE
            ' Accès impossible à un fichier non ouvert
            Err.Raise -290, "AdoDB.EnregistrerSous", "Le fichier n'a pas été ouvert."
        Case Trim(pNomFichier) = ""
            Err.Raise -291, "AdoDB.EnregistrerSous", "Le nom du fichier n'est pas défini."
        Case pTypeFichier = NON_DEFINI
            Err.Raise -292, "AdoDB.EnregistrerSous", "Le type de fichier n'est pas défini."
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

'-------------------------------------------------------------------------------------------------------------------------
'  Fermer un objet ADODB
'-------------------------------------------------------------------------------------------------------------------------
Public Sub Fermer()
    
    If pStatutFichier <> OPENED_FILE Then
        ' Accès impossible à un fichier non ouvert
        Err.Raise -300, "AdoDB.Fermer", "Le fichier n'a pas été ouvert."
    Else
        pFichierADODB.Close
        pStatutFichier = CLOSED_FILE
    End If
    
End Sub

'-------------------------------------------------------------------------------------------------------------------------
'  Compteurs
'-------------------------------------------------------------------------------------------------------------------------

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

'-------------------------------------------------------------------------------------------------------------------------
' Affiche la boite de dialogue qui permet de sélectionner le fichier afin d'enregistrer sous
'-------------------------------------------------------------------------------------------------------------------------
' Propriétés utilisées :
'   Filtre              (Facultatif)
'   TitreBoiteDialogue  (Facultatif)
'-------------------------------------------------------------------------------------------------------------------------
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

'-------------------------------------------------------------------------------------------------------------------------
' Affiche la boite de dialogue qui permet de sélectionner un fichier qui correspond aux filtres
'-------------------------------------------------------------------------------------------------------------------------
' Propriétés utilisées :
'   LibelleFiltre       (Facultatif)
'   ExtensionFiltre     (Facultatif)
'   TitreBoiteDialogue  (Facultatif)
'-------------------------------------------------------------------------------------------------------------------------
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

'-------------------------------------------------------------------------------------------------------------------------
' Affiche la boite de dialogue qui permet de sélectionner un répertoire
'-------------------------------------------------------------------------------------------------------------------------
' Propriété utilisée :
'   TitreBoiteDialogue  (Facultatif)
'-------------------------------------------------------------------------------------------------------------------------
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

'-------------------------------------------------------------------------------------------------------------------------
' Vérifie si un fichier est trop volumineux pour être traité par ADODB
'-------------------------------------------------------------------------------------------------------------------------
' Propriété utilisée :
'   NomFichier          (Obligatoire)
'-------------------------------------------------------------------------------------------------------------------------
Public Function EstFichierVolumineux(Optional NomFichier As String) As Boolean
    
    Dim NomFichierLocal As String
    
    Select Case True
        Case Trim(NomFichier) <> ""
            NomFichierLocal = NomFichier
        Case pNomFichier <> ""
            NomFichierLocal = pNomFichier
        Case Else
            Err.Raise -310, "AdoDB.EstFichierVolumineux", "Le nom du fichier n'est pas défini."
    End Select
    
    If LongueurFichier(NomFichierLocal) > 4400000000# Then
        EstFichierVolumineux = True
    Else
        EstFichierVolumineux = False
    End If

End Function

'-------------------------------------------------------------------------------------------------------------------------
' Retourne la longueur d'un fichier (en octets)
'-------------------------------------------------------------------------------------------------------------------------
' Paramètre utilisé :
'   NomFichier         (Facultatif)
' Propriété utilisée :
'   NomFichier          (Facultatif)
'-------------------------------------------------------------------------------------------------------------------------
Public Function LongueurFichier(Optional NomFichier As String) As Double

    Dim NomFichierLocal As String
    Dim oFileSystem As Object, oFichier As Object
    
    Select Case True
        Case Trim(NomFichier) <> ""
            NomFichierLocal = NomFichier
        Case pNomFichier <> ""
            NomFichierLocal = pNomFichier
        Case Else
            Err.Raise -320, "AdoDB.LongueurFichier", "Le nom du fichier n'est pas défini."
    End Select

    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
    Set oFichier = oFileSystem.GetFile(NomFichierLocal)
    LongueurFichier = oFichier.Size

End Function

'-------------------------------------------------------------------------------------------------------------------------
' Vérifie si un fichier existe
'-------------------------------------------------------------------------------------------------------------------------
' Paramètre utilisé :
'   NomFichier         (Facultatif)
' Propriété utilisée :
'   NomFichier          (Facultatif)
'-------------------------------------------------------------------------------------------------------------------------
Public Function FichierExiste(Optional NomFichier As String) As Boolean
    
    Dim NomFichierLocal As String
    
    Select Case True
        Case Trim$(NomFichier) <> ""
            NomFichierLocal = NomFichier
        Case pNomFichier <> ""
            NomFichierLocal = pNomFichier
        Case Else
            Err.Raise -330, "AdoDB.FichierExiste", "Le nom du fichier n'est pas défini."
    End Select

    FichierExiste = Dir(NomFichierLocal, vbNormal) <> ""

End Function

'-------------------------------------------------------------------------------------------------------------------------
' Vérifie si un répertoire existe
'-------------------------------------------------------------------------------------------------------------------------
' Propriété utilisée :
'   Repertoire          (Obligatoire)
'-------------------------------------------------------------------------------------------------------------------------
Public Function RepertoireExiste(Optional Repertoire As String) As Boolean
    
    Dim NomRepertoire As String
    
    Select Case True
        Case Trim(Repertoire) <> ""
            NomRepertoire = Repertoire
        Case pRepertoire <> ""
            NomRepertoire = pRepertoire
        Case Else
            Err.Raise -340, "AdoDB.RepertoireExiste", "Le nom du répertoire n'est pas défini."
    End Select

    RepertoireExiste = Dir(NomRepertoire, vbDirectory) <> ""

End Function

'-------------------------------------------------------------------------------------------------------------------------
' Convertir un nom de chemin défini par une URL OneDrive ou SharePoint vers un nom de chemin Windows
' Exemple : https://xxx-my.sharepoint.com/personal/ devient c:\Users\xxxx\OneDrive - xxx
'-------------------------------------------------------------------------------------------------------------------------
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

'-------------------------------------------------------------------------------------------------------------------------
' Retourne l'extension du fichier
'-------------------------------------------------------------------------------------------------------------------------
Public Function Extension() As String
    
    Dim lPosPt As Long
    
    lPosPt = InStrRev(pNomFichier, ".")
    If lPosPt > 0 Then
        Extension = LCase$(Right$(pNomFichier, Len(pNomFichier) - lPosPt))
    Else
        Extension = ""
    End If
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Initialisation de la classe : création des objets ADODB.Stream et le dictionnaire des encodages   *
'-------------------------------------------------------------------------------------------------------------------------
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
    DicEncodage.Add AD_CP_1250, "Windows-1250"
    DicEncodage.Add AD_CP_1251, "Windows-1251"
    DicEncodage.Add AD_CP_1252, "windows-1252"
    DicEncodage.Add AD_ISO_8859_11, "windows-874"
    DicEncodage.Add AD_ARABE, "iso-8859-6"
    DicEncodage.Add AD_AUTRE, "Autre"
    
    ' Fichier pas encore ouvert
    pStatutFichier = UNDEFINED_FILE_STATUS
    pNbreEnregLus = 0
    pNbreEnregEcrits = 0
    pNbreOctectsLus = 0
    pNbreOctectsEcrits = 0
    
End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Terminaison de la classe : libérer la mémoire en détruisant les objets
'-------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()

    Set pFichierADODB = Nothing
    Set DicEncodage = Nothing

End Sub
