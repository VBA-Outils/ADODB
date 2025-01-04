VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADODB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' VBA-ADODB v1.0
' https://github.com/VBA-Outils/ADODB
'
' Gestion des fichiers via ADODB
'
' @Class ADODB
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
' vbObjectError + 1000   : Erreur lors de l'appel d'une propriété ou méthode. Voir le message d'erreur.

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
'
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

' Proprieté "TypeFichier" : type de format du fichier
Public Enum EnumTypeFichier
    FICHIER_NON_DEFINI = 0
    FICHIER_BINAIRE = 1    ' Fichier Binaire
    FICHIER_TEXTE = 2      ' Fichier Texte
End Enum

' Propriété "SeparateurLigne" : caractère(s) de fin d'enregistrement
Public Enum EnumSeparateurLigne
    SEPARATEUR_NON_DEFINI = 0
    SEPARATEUR_CRLF = -1        ' Par défaut, Retour Chariot et Ligne Suivante
    SEPARATEUR_LF = 10          ' Ligne Suivante
    SEPARATEUR_CR = 13          ' Retour Chariot
End Enum

' Propriété "encodage" : encodage du fichier
' La liste des encodages n'est pas exhaustive et doit être complétée si nécessaire. Les plus utilisés sont présents.
Public Enum EnumListeEncodage
    NON_DEFINI = 0
    ASCII = 1
    LATIN_2 = 2
    LATIN_4 = 3
    CYRILLIQUE = 4
    GREC = 5
    LATIN_5 = 6
    COREEN = 7
    UTF_7 = 8
    UTF_8 = 9
    CP1250 = 10
    CP1251 = 11
    X_ANSI = 12
    ISO_8859_11 = 13
    ARABE = 14
    AUTRE = 15
End Enum

' Propriété "TypeAcces"
Public Enum EnumTypeAcces
    ACCES_NON_DEFINI = 0
    ACCES_LECTURE = 1           ' Ouverture en lecture du fichier
    ACCES_ECRITURE = 2          ' Ouverture en écriture du fichier
End Enum

' Filtres pour la boîte de dialogue "Enregistrer sous" (les plus utilisés)
Public Enum EnumIndexFiltre
    INDEX_FILTRE_NON_DEFINI = 0
    INDEX_FILTRE_XLSX = 1            ' Classeur Excel
    INDEX_FILTRE_XLSM = 2            ' Classeur Excel prenant en charge les macros
    INDEX_FILTRE_XLSB = 3            ' Classeur Excel binaire
    INDEX_FILTRE_XLS = 4             ' Classeur Excel 97-2003
    INDEX_FILTRE_CSV_COMA = 5        ' CSV UTF-8 (délimité par des virgules)
    INDEX_FILTRE_XML = 6             ' Données XML
    INDEX_FILTRE_XLTX = 9            ' Modèle Excel
    INDEX_FILTRE_XLTM = 10           ' Modèle Excel prenant en charge les macros
    INDEX_FILTRE_CSV_TAB = 12        ' CSV (séparateur : tabulation)
    INDEX_FILTRE_CSV_SEMICOLON = 16  ' CSV (séparateur : point-virgule)
End Enum

' Propriété de la méthode "Lire"
Public Enum EnumMethodeLecture
    LIRE_FICHIER = -1           ' Lit l'intégralité du fichier. C’est la seule valeur valide s'il s'agit d'un fichier binaire.
    LIRE_ENREG_SUIVANT = -2     ' Lit l'enregistrement suivant (jusqu'au prochain séparateur de lignes ou la fin de fichier).
End Enum

' Statut du flux
Private Enum EnumStatutFlux
    STATUT_FICHIER_INITIAL = 0
    STATUT_FICHIER_OUVERT = 1
    STATUT_FICHIER_FERME = 2
End Enum

' *---------------------------------------------------------------------------------------------------*
' * Propriétés des fichiers                                                                           *
' *---------------------------------------------------------------------------------------------------*

' Encodage du fichier texte
Private psEncodage As String
Private plEncodage As EnumListeEncodage
Private DicEncodage As Dictionary
' Séparateur de lignes (fichier texte uniquement)
Private plSeparateurLigne As EnumSeparateurLigne
' Type de fichier : Texte ou Binaire
Private plTypeFichier As EnumTypeFichier
' Nom physique du fichier (le répertoire peut être inclus dans le nom)
Private psNomFichier As String
' Objet ADODB.stream
Private poFichier As Object
' Type d'accès au fichier : lecture, écriture
Private psTypeAcces As EnumTypeAcces

' *---------------------------------------------------------------------------------------------------*
' * Gestion des boîtes de dialogue                                                                    *
' *---------------------------------------------------------------------------------------------------*

' Titre de la boite de dialogue (fichiers)
Private psTitreBoiteDialogue As String
' Libellé du filtre sur les fichiers à appliquer
Private psLibelleFiltre As String
' Extensions appliquées dans le filtre
Private psExtensionFiltre As String
' Nom du fichier initial proposé dans la boite de dialogue Enregistrer sous
Private psNomInitialFichier As String
' Répertoire sélectionné via la boite de dialogue
Private psRepertoire As String
' Filtre pour la boîte à dialogue d'enregsitrement sous
Private plIndexFiltre As EnumIndexFiltre

' *---------------------------------------------------------------------------------------------------*
' * Gestion des fichiers                                                                              *
' *---------------------------------------------------------------------------------------------------*

' Fichier ouvert
Private plStatutFlux As EnumStatutFlux
Private plNbreEnregLus As Long
Private plNbreEnregEcrits As Long
Private plNbreOctectsLus As Long
Private plNbreOctectsEcrits As Long

' *---------------------------------------------------------------------------------------------------*
' * Propriété : objet fichier                                                                         *
' *---------------------------------------------------------------------------------------------------*
Property Set Fichier(oFichier As Object)
    Set poFichier = oFichier
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : type de fichier (texte, binaire)                                                      *
' *---------------------------------------------------------------------------------------------------*
Property Let TypeFichier(lTypeFichier As EnumTypeFichier)

    ' Vérifier le type de fichier
    Select Case lTypeFichier
        Case FICHIER_BINAIRE:
            Select Case True
                Case psEncodage <> "":
                    ' L'encodage n'est pas nécessaire pour lire un fichier binaire
                    Err.Raise vbObjectError + 1000, , "[TypeFichier] L'encodage ne doit pas être défini pour les fichiers binaires."
                Case plSeparateurLigne <> SEPARATEUR_NON_DEFINI:
                    ' Le séparateur de lignes n'est pas nécessaire pour lire un fichier binaire
                    Err.Raise vbObjectError + 1000, , "[TypeFichier] Le séparateur de lignes ne doit pas être défini pour les fichiers binaires."
                Case Else:
                    plTypeFichier = lTypeFichier
            End Select
        Case FICHIER_TEXTE:
            plTypeFichier = lTypeFichier
        Case Else
            Err.Raise vbObjectError + 1000, , "[TypeFichier] Le type de fichier " & lTypeFichier & " est invalide."
    End Select

End Property

Property Get TypeFichier() As EnumTypeFichier
    TypeFichier = plTypeFichier
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : Encodage du fichier texte (Enum)                                                      *
' *---------------------------------------------------------------------------------------------------*
Property Let Encodage(lEncodage As EnumListeEncodage)

    Select Case True
        ' L'encodage doit être renseigné
        Case lEncodage = NON_DEFINI:
            Err.Raise vbObjectError + 1000, , "[Encodage] L'encodage n'est pas renseigné."
        Case plTypeFichier = FICHIER_BINAIRE:
            ' L'encodage n'est pas nécessaire pour lire un fichier binaire
            Err.Raise vbObjectError + 1000, , "[Encodage] L'encodage ne doit pas être défini pour les fichiers binaires."
            ' Vérifie que l'encodage existe dans l'enum
        Case EnumEncodageExiste(lEncodage) = True:
            plEncodage = lEncodage
            psEncodage = DicEncodage.Keys(lEncodage)
        Case Else:
            ' Si l'encodage est absent de l'enum ListeEncodage et du dictionnaire dicEncodage alors une erreur est retournée.
            ' La liste des encodages n'est pas exhaustive et doit être complétée si nécessaire, sinon utiliser la propriété EncodageTxt.
            Err.Raise vbObjectError + 2, , "[Encodage] L'encodage n'existe pas dans l'enum EnumListeEncodage."
    End Select
    
End Property

Property Get Encodage() As EnumListeEncodage
    
    Encodage = plEncodage
    
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : Encodage du fichier texte (code technique)                                            *
' *---------------------------------------------------------------------------------------------------*
Property Let EncodageTxt(sEncodage As String)

    Select Case True
        ' L'encodage doit être renseigné
        Case Trim(sEncodage) = "":
            Err.Raise vbObjectError + 3, , "[EncodageTxt] L'encodage n'est pas renseigné."
        ' L'encodage n'est pas nécessaire pour lire un fichier binaire
        Case plTypeFichier = FICHIER_BINAIRE:
            Err.Raise vbObjectError + 1000, , "[EncodageTxt] L'encodage ne doit pas être défini pour les fichiers binaires."
        ' Aucun contrôle de l'encodage, la saisie est libre.
        Case Else:
            psEncodage = sEncodage
            plEncodage = AUTRE
    End Select
    
End Property

Property Get EncodageTxt() As String
    EncodageTxt = psEncodage
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : séparateur de lignes (CL, LF, CR/LF)                                                  *
' *---------------------------------------------------------------------------------------------------*
Property Let SeparateurLigne(lSeparateurLigne As EnumSeparateurLigne)

    If lSeparateurLigne = SEPARATEUR_CR Or lSeparateurLigne = SEPARATEUR_CRLF Or lSeparateurLigne = SEPARATEUR_LF Then
        If plTypeFichier = FICHIER_BINAIRE Then
            ' Le séparateur ne sert à rien pour les fichiers binaires
            Err.Raise vbObjectError + 1000, , "[SeparateurLigne] Le séparateur ne doit pas être défini pour les fichiers binaires."
        Else
            plSeparateurLigne = lSeparateurLigne
        End If
    Else
        Err.Raise vbObjectError + 1000, , "[SeparateurLigne] Le séparateur " & lSeparateurLigne & " est invalide."
    End If
    
End Property

Property Get SeparateurLigne() As EnumSeparateurLigne
    SeparateurLigne = plSeparateurLigne
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : type d'accès au fichier                                                               *
' *---------------------------------------------------------------------------------------------------*
Property Let TypeAcces(sTypeAcces As EnumTypeAcces)

    If sTypeAcces = ACCES_ECRITURE Or sTypeAcces = ACCES_LECTURE Then
        psTypeAcces = sTypeAcces
    Else
        Err.Raise vbObjectError + 1000, , "[TypeAcces] Le type d'accès " & sTypeAcces & " est invalide."
    End If
    
End Property

Property Get TypeAcces() As EnumTypeAcces
    TypeAcces = psTypeAcces
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : nom du fichier                                                                        *
' *---------------------------------------------------------------------------------------------------*
Property Let NomFichier(sNomFichier As String)

    If Trim(sNomFichier) = "" Then
        Err.Raise vbObjectError + 1000, , "[NomFichier] Le nom du fichier doit être renseigné."
    Else
        ' Convertir les noms de fichier sous forme d'URL au format Lecteur:\Répertoire\NomFichier.extension
        psNomFichier = ConvertirUrlSharePoint(sNomFichier)
    End If
    
End Property

Property Get NomFichier() As String
    NomFichier = psNomFichier
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : titre de la boîte de dialogue à afficher                                              *
' *---------------------------------------------------------------------------------------------------*
Property Let TitreBoiteDeDialogue(sTitre As String)

    If Trim(sTitre) = "" Then
        Err.Raise vbObjectError + 1000, , "[TitreBoiteDeDialogue] Le titre de la boîte de dialogue doit être renseigné."
    Else
        psTitreBoiteDialogue = sTitre
    End If
    
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : libellé du filtre à appliquer afin de sélectionner les fichiers                       *
' *---------------------------------------------------------------------------------------------------*
Property Let LibelleFiltre(sLibelleFiltre As String)

    If Trim(sLibelleFiltre) = "" Then
        Err.Raise vbObjectError + 1000, , "[LibelleFiltre] Le libellé du filtre doit être renseigné."
    Else
        psLibelleFiltre = sLibelleFiltre
    End If
    
End Property

Property Get LibelleFiltre() As String
    LibelleFiltre = psLibelleFiltre
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : Filtre des extensions séparées par une virgule, eg. *.csv,*.txt                       *
' *---------------------------------------------------------------------------------------------------*
Property Let ExtensionFiltre(sExtensionFiltre As String)

    If Trim(sExtensionFiltre) = "" Then
        Err.Raise vbObjectError + 1000, , "[ExtensionFiltre] Les extensions doivent être renseignées."
    Else
        psExtensionFiltre = sExtensionFiltre
    End If
    
End Property

Property Get ExtensionFiltre() As String
    ExtensionFiltre = psExtensionFiltre
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : nom du fichier pré renseigné dans la boite de dialogue pour enregistrer sous          *
' *---------------------------------------------------------------------------------------------------*
Property Let NomInitialFichier(sNomInitialFichier As String)

    If Trim(sNomInitialFichier) = "" Then
        Err.Raise vbObjectError + 1000, , "[NomInitialFichier] Le nom initial du fichier n'est pas renseigné."
    Else
        psNomInitialFichier = sNomInitialFichier
    End If

End Property

Property Get NomInitialFichier() As String
    NomInitialFichier = psNomInitialFichier
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : chemin d'un répertoire (dossier)                                                      *
' *---------------------------------------------------------------------------------------------------*
Property Let Repertoire(sRepertoire As String)
    
    If Trim(sRepertoire) = "" Then
        Err.Raise vbObjectError + 1000, , "[Repertoire] Le répertoire n'est pas renseigné."
    End If
    
    psRepertoire = sRepertoire
    
End Property

Property Get Repertoire() As String
    Repertoire = psRepertoire
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : Filtre à appliquer lors de la sélection des fichiers à enregistrer sous               *
' *---------------------------------------------------------------------------------------------------*
Property Let Filtre(lIndexFiltre As EnumIndexFiltre)
    
    If lIndexFiltre < 1 Or lIndexFiltre > 29 Then
        Err.Raise vbObjectError + 1000, , "[lIndexFiltre] Le filtre " & lIndexFiltre & " est invalide."
    Else
        plIndexFiltre = lIndexFiltre
    End If
    
End Property

Property Get Filtre() As EnumIndexFiltre
    Filtre = plIndexFiltre
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
        Case plTypeFichier = FICHIER_NON_DEFINI:
            Err.Raise vbObjectError + 1000, , "[Ouvrir] Le type de fichier n'est pas défini."
        ' Vérifier que l'encodage est renseigné si le fichier est de type Texte
        plTypeFichier = FICHIER_TEXTE And psEncodage = "":
            Err.Raise vbObjectError + 1000, , "[Ouvrir] L'encodage n'est pas défini."
        ' Vérifier que le type d'accès est renseigné
        Case psTypeAcces = ACCES_NON_DEFINI:
            Err.Raise vbObjectError + 1000, , "[Ouvrir] Le type d'accès n'est pas défini."
        Case psTypeAcces = ACCES_LECTURE:
            ' Si le nom de fichier n'est pas renseigné
            If Trim(psNomFichier) = "" Then
                Err.Raise vbObjectError + 1000, , "[Ouvrir] Le nom du fichier n'est pas défini."
            End If
            ' Vérifier que le fichier existe si l'accès est en lecture
            If FichierExiste = False Then
                Err.Raise vbObjectError + 1000, , "[Ouvrir] Le fichier " & psNomFichier & " n'existe pas."
            End If
    End Select
    
    With poFichier
        ' Type de fichier : binaire ou texte
        .Type = plTypeFichier
        ' Si fichier texte, préciser l'encodage et le séparateur
        If plTypeFichier = FICHIER_TEXTE Then
            .Charset = psEncodage
            If plSeparateurLigne <> SEPARATEUR_NON_DEFINI Then
                .LineSeparator = plSeparateurLigne
            End If
        End If
        .Open
        ' Si ouverture en lecture alors on charge le fichier
        If psTypeAcces = ACCES_LECTURE Then
            .LoadFromFile (psNomFichier)
        End If
    End With
    ' Fichier ouvert sans erreur
    plStatutFlux = STATUT_FICHIER_OUVERT
    
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
        Case plStatutFlux <> STATUT_FICHIER_OUVERT:
            Err.Raise vbObjectError + 1000, , "[Lire] Le fichier n'a pas été ouvert."
        ' Le nombre de caractères à lire doit être > 0
        Case lNbreCarALire <= 0:
            Err.Raise vbObjectError + 1000, , "[Lire] Le nombre de caractères à lire n'est pas valide."
        ' Seule la lecture complète d'un fichier binaire est possible
        Case plTypeFichier = FICHIER_BINAIRE:
            Err.Raise vbObjectError + 1000, , "[Lire] La lecture partielle d'un fichier binaire n'est pas possible."
        Case Else:
            ' Lire les n caractères suivants
            Lire = poFichier.ReadText(lNbreCarALire)
            plNbreOctectsLus = plNbreOctectsLus + Len(Lire)
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
        Case plStatutFlux <> STATUT_FICHIER_OUVERT:
            Err.Raise vbObjectError + 1000, , "[LireEnregistrement] Le fichier n'a pas été ouvert."
        ' Vérifie que le séparateur a bien été défini pour la lecture par ligne
        Case plSeparateurLigne = SEPARATEUR_NON_DEFINI:
            Err.Raise vbObjectError + 1000, , "[LireEnregistrement] Le séparateur de ligne n'est pas défini."
        ' Lecture du prochain enregistrement
        Case plTypeFichier = FICHIER_TEXTE:
            LireEnregistrement = poFichier.ReadText(LIRE_ENREG_SUIVANT)
            plNbreEnregLus = plNbreEnregLus + 1
            plNbreOctectsLus = plNbreOctectsLus + Len(LireEnregistrement)
        ' Les fichiers binaires ne comportent pas de séparateur, par conséquent, il n'est pas possible de lire un enregistrement
        Case plTypeFichier = FICHIER_BINAIRE:
            Err.Raise vbObjectError + 1000, , "[LireEnregistrement] La lecture d'un enregistrement d'un fichier binaire n'est pas possible."
    End Select
    
End Function

' *---------------------------------------------------------------------------------------------------*
' * Fin de fichier atteinte ?                                                                         *
' *---------------------------------------------------------------------------------------------------*
Public Function FinFichier() As Boolean
    
    ' Accès impossible à un fichier non ouvert
    If plStatutFlux <> STATUT_FICHIER_OUVERT Then
        Err.Raise vbObjectError + 1000, , "[FinFichier] Le fichier n'a pas été ouvert."
    End If
    
    FinFichier = poFichier.EOS

End Function

' *---------------------------------------------------------------------------------------------------*
' * Lecture intégrale d'un fichier ADODB                                                              *
' *---------------------------------------------------------------------------------------------------*
Public Function LireFichier() As Variant
    
    Select Case True
        ' Accès impossible à un fichier non ouvert
        Case plStatutFlux <> STATUT_FICHIER_OUVERT:
            Err.Raise vbObjectError + 1000, , "[LireFichier] Le fichier n'a pas été ouvert."
        ' Lire entièrement un fichier binaire
        Case plTypeFichier = FICHIER_BINAIRE:
            LireFichier = poFichier.Read(LIRE_FICHIER)
            plNbreOctectsLus = plNbreOctectsLus + Len(LireFichier)
        ' Lire entièrement un fichier texte
        Case plTypeFichier = FICHIER_TEXTE:
            LireFichier = poFichier.ReadText(LIRE_FICHIER)
            plNbreOctectsLus = plNbreOctectsLus + Len(LireFichier)
    End Select

End Function

' *---------------------------------------------------------------------------------------------------*
' * Ecrire un enregistrement dans un flux ADODB.stream                                                *
' *---------------------------------------------------------------------------------------------------*
' * Paramètres :                                                                                      *
' *   -- sEnregistrement (Obligatoire) : enregistrement à écrire (XML, json, texte, etc)              *
' *---------------------------------------------------------------------------------------------------*
Public Sub EcrireEnregistrement(sEnregistrement As String)
    
    Select Case True
        ' Accès impossible à un fichier non ouvert
        Case plStatutFlux <> STATUT_FICHIER_OUVERT:
            Err.Raise vbObjectError + 26, , "[EcrireEnregistrement] Le fichier n'a pas été ouvert."
        ' Ecriture de l'enregistrement dans le fichier texte. Les séparateurs CR/LF doivent être présents dans l'enregistrement.
        Case plTypeFichier = FICHIER_TEXTE:
            poFichier.WriteText sEnregistrement
            plNbreEnregEcrits = plNbreEnregEcrits + 1
            plNbreOctectsEcrits = plNbreOctectsEcrits + Len(sEnregistrement)
        ' Il n'est pas possible d'écrire un enregistrement dans un fichier binaire.
        Case plTypeFichier = FICHIER_BINAIRE:
            Err.Raise vbObjectError + 1000, , "[EcrireEnregistrement] L'écriture d'un enregistrement dans un fichier binaire n'est pas possible."
    End Select
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Ecrire des données dans un flux ADODB.stream                                                      *
' *---------------------------------------------------------------------------------------------------*
' * Paramètres :                                                                                      *
' *   -- vDonnees (Obligatoire) : données à écrire (image, pdf, etc)                                  *
' *---------------------------------------------------------------------------------------------------*
Public Sub Ecrire(vDonnees As Variant)
    
    Select Case True
        ' Accès impossible à un fichier non ouvert
        Case plStatutFlux <> STATUT_FICHIER_OUVERT:
            Err.Raise vbObjectError + 26, , "[Ecrire] Le fichier n'a pas été ouvert."
        ' Ecrire le fichier binaire (en une seule fois)
        Case plTypeFichier = FICHIER_BINAIRE:
            poFichier.Write vDonnees
            plNbreOctectsEcrits = plNbreOctectsEcrits + Len(vDonnees)
        ' Il n'est pas possible d'écrire des données brutes dans un fichier texte.
        Case plTypeFichier = FICHIER_TEXTE:
            Err.Raise vbObjectError + 1000, , "[Ecrire] L'écriture de données binaires dans fichier texte n'est pas possible."
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
        Case plStatutFlux <> STATUT_FICHIER_OUVERT:
            ' Accès impossible à un fichier non ouvert
            Err.Raise vbObjectError + 26, , "[EnregistrerSous] Le fichier n'a pas été ouvert."
        Case Trim(psNomFichier) = "":
            Err.Raise vbObjectError + 1000, , "[EnregistrerSous] Le nom du fichier n'est pas défini."
        Case plTypeFichier = FICHIER_NON_DEFINI:
            Err.Raise vbObjectError + 1000, , "[EnregistrerSous] Le type de fichier n'est pas défini."
        Case Else
            poFichier.SaveToFile psNomFichier, plTypeFichier
    End Select
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Fermer un objet ADODB                                                                             *
' *---------------------------------------------------------------------------------------------------*
Public Sub Fermer()
    
    If plStatutFlux <> STATUT_FICHIER_OUVERT Then
        ' Accès impossible à un fichier non ouvert
        Err.Raise vbObjectError + 26, , "[Fermer] Le fichier n'a pas été ouvert."
    Else
        poFichier.Close
        plStatutFlux = STATUT_FICHIER_FERME
    End If
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Compteurs                                                                                         *
' *---------------------------------------------------------------------------------------------------*

Public Function NbreEnregLus() As Long
    NbreEnregLus = plNbreEnregLus
End Function

Public Function NbreEnregEcrits() As Long
    NbreEnregEcrits = plNbreEnregEcrits
End Function

Public Function NbreOctectsLus() As Long
    NbreOctectsLus = plNbreOctectsLus
End Function

Public Function NbreOctectsEcrits() As Long
    NbreOctectsEcrits = plNbreOctectsEcrits
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
        If psNomInitialFichier <> "" Then
            .InitialFileName = psNomInitialFichier
        End If
        .InitialView = msoFileDialogViewDetails
        If plIndexFiltre <> INDEX_FILTRE_NON_DEFINI Then
            .FilterIndex = plIndexFiltre
        End If
        ' Titre par défaut de la boîte de dialogue
        If psTitreBoiteDialogue = "" Then
            .Title = "Sélectionner le répertoire et saisir le nom du fichier à enregistrer sous"
        Else
            .Title = psTitreBoiteDialogue
        End If
        If .Show = -1 Then
            psNomFichier = .SelectedItems(1)
        Else
            psNomFichier = ""
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
        If psNomInitialFichier <> "" Then
            .InitialFileName = psNomInitialFichier
        End If
        .InitialView = msoFileDialogViewDetails
        ' Si filtre ou extension non renseigné
        If psLibelleFiltre = "" Or psExtensionFiltre = "" Then
            .Filters.Add "Tous les fichiers", "*.*", 1
        Else
            .Filters.Add psLibelleFiltre, psExtensionFiltre, 1
        End If
        .FilterIndex = 1
        ' Titre par défaut si non renseigné
        If psTitreBoiteDialogue = "" Then
            .Title = "Sélectionner le répertoire et le fichier"
        Else
            .Title = psTitreBoiteDialogue
        End If
        If .Show = -1 Then
            psNomFichier = .SelectedItems(1)
        Else
            psNomFichier = ""
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
        If psNomInitialFichier <> "" Then
            .InitialFileName = psNomInitialFichier
        End If
        .InitialView = msoFileDialogViewDetails
        If psTitreBoiteDialogue = "" Then
            .Title = "Sélectionner le répertoire"
        Else
            .Title = psTitreBoiteDialogue
        End If
        ' Si aucun répertoire sélectionné alors fin du traitement
        If .Show = -1 Then
            psRepertoire = .SelectedItems(1)
        Else
            psRepertoire = ""
        End If
    End With

End Sub

' *---------------------------------------------------------------------------------------------------*
' * Vérifie si un fichier est trop volumineux pour être traité par ADODB                              *
' *---------------------------------------------------------------------------------------------------*
' * Propriété utilisée :                                                                              *
' *   -- NomFichier          (Obligatoire)                                                            *
' *---------------------------------------------------------------------------------------------------*
Public Function EstFichierVolumineux(Optional sNomFichier As String) As Boolean
    
    Dim sNomFichierLocal As String
    
    Select Case True
        Case Trim(sNomFichier) <> "": sNomFichierLocal = sNomFichier
        Case psNomFichier <> "": sNomFichierLocal = psNomFichier
        Case Else: Err.Raise vbObjectError + 1000, , "EstFichierVolumineux : Le nom du fichier n'est pas défini."
    End Select
    
    If LongueurFichier(sNomFichierLocal) > 4400000000# Then
        EstFichierVolumineux = True
    Else
        EstFichierVolumineux = False
    End If

End Function

' *---------------------------------------------------------------------------------------------------*
' * Retourne la longueur d'un fichier (en octets)                                                     *
' *---------------------------------------------------------------------------------------------------*
' * Paramètre utilisé :                                                                               *
' *   -- sNomFichier         (Facultatif)                                                             *
' * Propriété utilisée :                                                                              *
' *   -- NomFichier          (Facultatif)                                                             *
' *---------------------------------------------------------------------------------------------------*
Public Function LongueurFichier(Optional sNomFichier As String) As Double

    Dim sNomFichierLocal As String
    Dim oFileSystem As Object, oFichier As Object
    
    Select Case True
        Case Trim(sNomFichier) <> "": sNomFichierLocal = sNomFichier
        Case psNomFichier <> "": sNomFichierLocal = psNomFichier
        Case Else: Err.Raise vbObjectError + 1000, , "LongueurFichier : Le nom du fichier n'est pas défini."
    End Select

    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
    Set oFichier = oFileSystem.GetFile(sNomFichierLocal)
    LongueurFichier = oFichier.Size

End Function

' *---------------------------------------------------------------------------------------------------*
' * Vérifie si un fichier existe                                                                      *
' *---------------------------------------------------------------------------------------------------*
' * Paramètre utilisé :                                                                               *
' *   -- sNomFichier         (Facultatif)                                                             *
' * Propriété utilisée :                                                                              *
' *   -- NomFichier          (Facultatif)                                                             *
' *---------------------------------------------------------------------------------------------------*
Public Function FichierExiste(Optional sNomFichier As String) As Boolean
    
    Dim sNomFichierLocal As String
    
    Select Case True
        Case Trim(sNomFichier) <> "": sNomFichierLocal = sNomFichier
        Case psNomFichier <> "": sNomFichierLocal = psNomFichier
        Case Else: Err.Raise vbObjectError + 1000, , "FichierExiste : le nom du fichier n'est pas défini."
    End Select

    FichierExiste = Dir(sNomFichierLocal, vbNormal) <> ""

End Function

' *---------------------------------------------------------------------------------------------------*
' * Vérifie si un répertoire existe                                                                   *
' *---------------------------------------------------------------------------------------------------*
' * Propriété utilisée :                                                                              *
' *   -- Repertoire          (Obligatoire)                                                            *
' *---------------------------------------------------------------------------------------------------*
Public Function RepertoireExiste(Optional sRepertoire As String) As Boolean
    
    Dim sNomRepertoire As String
    
    Select Case True
        Case Trim(sRepertoire) <> "": sNomRepertoire = sRepertoire
        Case psRepertoire <> "": sNomRepertoire = psRepertoire
        Case Else: Err.Raise vbObjectError + 1000, , "RepertoireExiste : le nom du répertoire n'est pas défini."
    End Select

    RepertoireExiste = Dir(sNomRepertoire, vbDirectory) <> ""

End Function

' *---------------------------------------------------------------------------------------------------*
' * Vérifie si l'encodage existe dans le dictionnaire                                                 *
' *---------------------------------------------------------------------------------------------------*
Private Function EnumEncodageExiste(lEncodage As Long) As Boolean

    Dim iIcDico As Integer
    
    For iIcDico = 0 To DicEncodage.Count
        If DicEncodage.Items(iIcDico) = lEncodage Then
            EnumEncodageExiste = True
            Exit For
        End If
    Next iIcDico
    
End Function

' *--------------------------------------------------------------------------------------------------------------------------*
' * Convertir un nom de chemin défini par une URL OneDrive ou SharePoint vers un nom de chemin Windows                       *
' * Exemple : https://xxx-my.sharepoint.com/personal/ devient c:\Users\xxxx\OneDrive - xxx                                   *
' *--------------------------------------------------------------------------------------------------------------------------*
Private Function ConvertirUrlSharePoint(sChemin As String) As String

    Dim sListeDossiers() As String, iNbDossiers As Integer, lPosDoc As Long, sRepertoire As String
    
    ' Si le chemin du fichier commence par http
    If LCase$(Left(sChemin, 4)) = "http" Then
        Select Case True
        ' Espace personnel sur SharePoint (i.e. OneDrive Commercial) ?
        Case sChemin Like "https://*-my.sharepoint.com/personal/*":
            ' Recherche la chaîne "/Documents/documents" afin d'obtenir le début de l'arborescence dans le dossier des documents
            lPosDoc = InStr(1, sChemin, "/Documents/Documents/", vbTextCompare) + Len("/Documents")
            ' Le répertoire local est récupéré à partir du 2ème /Documents
            sRepertoire = Mid(sChemin, lPosDoc, Len(sChemin) - lPosDoc + 1)
            ConvertirUrlSharePoint = Environ("OneDriveCommercial") & Replace(sRepertoire, "/", "\")
        ' Espace de travail partagé
        Case sChemin Like "https://weshare*":
            sListeDossiers = Split(sChemin, "/")
            ConvertirUrlSharePoint = "\\" & sListeDossiers(2) & "@SSL\DavWWWRoot"
            For iNbDossiers = 3 To UBound(sListeDossiers)
                ConvertirUrlSharePoint = ConvertirUrlSharePoint & "\" & sListeDossiers(iNbDossiers)
            Next
        Case sChemin Like "https://d.docs.live.net/*":
            ' Recherche la chaîne "/documents" afin d'obtenir le début de l'arborescence dans le dossier des documents
            lPosDoc = InStr(1, sChemin, "/Documents/", vbTextCompare)
            ' Le répertoire local est récupéré à partir du 2ème /Documents
            sRepertoire = Mid(sChemin, lPosDoc, Len(sChemin) - lPosDoc + 1)
            ConvertirUrlSharePoint = Environ("OneDrive") & Replace(sRepertoire, "/", "\")
        End Select
    Else
        ConvertirUrlSharePoint = sChemin
    End If
    
End Function

' *---------------------------------------------------------------------------------------------------*
' * Initialisation de la classe : création des objets ADODB.Stream et le dictionnaire des encodages   *
' *---------------------------------------------------------------------------------------------------*
Private Sub Class_Initialize()
    
    ' Création d'un objet ADODB
    Set poFichier = CreateObject("ADODB.STREAM")

    ' Création d'un dictionnaire des encodages
    Set DicEncodage = New Dictionary
    ' Ajout des encodages (à compléter si nécessaire)
    ' clé : libellé de l'encodage identique aux clés de registre charset de windows
    ' La liste ci-dessous doit être identique à l'enum EnumListeEncodage
    DicEncodage.Add "Non défini", 0       ' NON_DEFINI
    DicEncodage.Add "iso-8859-1", 1       ' ASCII
    DicEncodage.Add "iso-8859-2", 2       ' LATIN_2
    DicEncodage.Add "iso-8859-4", 3       ' LATIN_4
    DicEncodage.Add "iso-8859-5", 4       ' CYRILLIQUE
    DicEncodage.Add "iso-8859-7", 5       ' GREC
    DicEncodage.Add "iso-8859-9", 6       ' LATIN_5
    DicEncodage.Add "ks_c_5601-1987", 7   ' COREEN
    DicEncodage.Add "utf-7", 8            ' UTF_7
    DicEncodage.Add "utf-8", 9            ' UTF_8
    DicEncodage.Add "Windows-1250", 10    ' CP1250
    DicEncodage.Add "Windows-1251", 11    ' CP1251
    DicEncodage.Add "windows-1252", 12    ' X_ANSI
    DicEncodage.Add "windows-874", 13     ' ISO_8859_11
    DicEncodage.Add "iso-8859-6", 14      ' ARABE
    DicEncodage.Add "Autre", 15           ' AUTRE

    ' Fichier pas encore ouvert
    plStatutFlux = STATUT_FICHIER_INITIAL
    plNbreEnregLus = 0
    plNbreEnregEcrits = 0
    plNbreOctectsLus = 0
    plNbreOctectsEcrits = 0
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Terminaison de la classe : libérer la mémoire en détruisant les objets                            *
' *---------------------------------------------------------------------------------------------------*
Private Sub Class_Terminate()
    Set poFichier = Nothing
    Set DicEncodage = Nothing
End Sub
