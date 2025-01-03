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
' vbObjectError + 1000   : Erreur lors de l'appel d'une propriete ou methode. Voir le message d'erreur.

' Necessite d'activer la reference "Microsoft Scripting RunTime"
'
' Dans l'editeur de macros (Alt+F11): Menu Outils \ References
' Cochez les lignes :
' -- "Microsoft Scripting RunTime"
' Cliquez sur le bouton OK pour valider.

Option Explicit

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'
' Proprietes de la classe :
' -------------------------
'   TypeFichier                        : Type du fichier (texte ou binaire)
'   Encodage                           : Encodage du fichier (enum)
'   EncodageTxt                        : Encodage du fichier (saisie libre)
'   SeparateurLigne                    : Separateur de lignes d'un fichier texte
'   TypeAcces                          : Acces au fichier (lecture, ecriture)
'   TitreBoiteDeDialogue               : Titre affiche dans les boites de dialogue
'   LibelleFiltre                      : Libelle du filtre utilise pour afficher les fichiers dans une boite de dialogue
'   ExtensionFiltre                    : Filtre des extensions de fichier utilise dans une boite de dialogue
'   NomInitialFichier                  : Nom initial du fichier affiche dans une boite de dialogue lors d'un enregsitrement sous
'   Filtre                             : Filtre des extensions (parmi celles proposees par Excel) de fichier utilise dans une boite de dialogue
'   NomFichier                         : nom du fichier (y compris le chemin d'acces)
'   Repertoire                         : Repertoire selectionne dans une boite de dialogue
'   Fichier                            : Permettre de pointer sur un objet ADODB Stream en dehors de la classe
'   NbreEnregLus                       : Nombre d'enregistrements lus
'   NbreEnregEcrits                    : Nombre d'enregistrements ecrits
'   NbreOctectsLus                     : Nombre d'octets lus
'   NbreOctectsEcrits                  : Nombre d'octets ecrits
'
' Methodes de la classe :
' -----------------------
'   Ouvrir                             : Ouvrir un fichier via ADODB
'   LireFichier                        : Lire l'integralite d'un fichier (texte ou binaire)
'   LireEnregistrement                 : Lire un fichier texte jusqu'au prochain separateur
'   FinFichier                         : Fin du fichier texte atteinte
'   Lire                               : Lire n caracteres d'un fichier (texte ou binaire)
'   LireEnregistrement                 : Lire le prochain enregistrement (jusqu'au prochain separateur de lignes) d'un fichier texte
'   LireFichier                        : Lire l'integralite d'un fichier texte ou binaire
'   Fermer                             : Fermer un fichier
'   Ecrire                             : Ecrire des donnees dans un fichier texte ou binaire
'   EcrireEnregistrement               : Ecrire un enregistrement dans un fichier texte
'   EnregistrerSous                    : Enregistrer sous le nom de fichier le contenu de l'objet ADODB.Stream
'   SelectionnerFichierEnregistrerSous : Afficher la boite de dialogue de selection d'un fichier a enregsitrer sous (avec saisie du nom du fichier)
'   SelectionnerFichier                : Afficher la boite de dialogue de selection d'un fichier
'   SelectionnerRepertoire             : Afficher la boite de dialogue de selection d'un repertoire
'   EstFichierVolumineux               : Verifier si un fichier depasse 4 Go (taille maximale des fichiers texte via ADODB.stream)
'   FichierExiste                      : Verifier si le fichier dont le nom est "NomFichier" existe
'   RepertoireExiste                   : Verifier si le repertoire dont le nom est "NomRepertoire" existe
'   LongueurFichier                    : Retourner la longueur d'un fichier en octets
'
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

' Propriete "TypeFichier" : type de format du fichier
Public Enum TypeFichier
    FICHIER_NON_DEFINI = 0
    FICHIER_BINAIRE = 1    ' Fichier Binaire
    FICHIER_TEXTE = 2      ' Fichier Texte
End Enum

' Propriete "SeparateurLigne" : caractere(s) de fin d'enregistrement
Public Enum SeparateurLigne
    SEPARATEUR_NON_DEFINI = 0
    SEPARATEUR_CRLF = -1        ' Par defaut, Retour Chariot et Ligne Suivante
    SEPARATEUR_LF = 10          ' Ligne Suivante
    SEPARATEUR_CR = 13          ' Retour Chariot
End Enum

' Propriete "encodage" : encodage du fichier
' La liste des encodages n'est pas exhaustive et doit etre completee si necessaire. Les plus utilises sont presents.
Public Enum ListeEncodage
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

' Propriete "TypeAcces"
Public Enum TypeAcces
    ACCES_NON_DEFINI = 0
    ACCES_LECTURE = 1           ' Ouverture en lecture du fichier
    ACCES_ECRITURE = 2          ' Ouverture en ecriture du fichier
End Enum

' Filtres pour la boite de dialogue "Enregistrer sous" (les plus utilises)
Public Enum IndexFiltre
    INDEX_FILTRE_NON_DEFINI = 0
    INDEX_FILTRE_XLSX = 1            ' Classeur Excel
    INDEX_FILTRE_XLSM = 2            ' Classeur Excel prenant en charge les macros
    INDEX_FILTRE_XLSB = 3            ' Classeur Excel binaire
    INDEX_FILTRE_XLS = 4             ' Classeur Excel 97-2003
    INDEX_FILTRE_CSV_COMA = 5        ' CSV UTF-8 (delimite par des virgules)
    INDEX_FILTRE_XML = 6             ' Donnees XML
    INDEX_FILTRE_XLTX = 9            ' Modele Excel
    INDEX_FILTRE_XLTM = 10           ' Modele Excel prenant en charge les macros
    INDEX_FILTRE_CSV_TAB = 12        ' CSV (separateur : tabulation)
    INDEX_FILTRE_CSV_SEMICOLON = 16  ' CSV (separateur : point-virgule)
End Enum

' Propriete de la methode "Lire"
Public Enum MethodeLecture
    LIRE_FICHIER = -1           ' Lit l'integralite du fichier. C'est la seule valeur valide s'il s'agit d'un fichier binaire.
    LIRE_ENREG_SUIVANT = -2     ' Lit l'enregistrement suivant (jusqu'au prochaine separateur de lignes ou la fin de fichier).
End Enum

' Statut du flux
Private Enum StatutFlux
    STATUT_FICHIER_INITIAL = 0
    STATUT_FICHIER_OUVERT = 1
    STATUT_FICHIER_FERME = 2
End Enum

' *---------------------------------------------------------------------------------------------------*
' * Proprietes des fichiers                                                                           *
' *---------------------------------------------------------------------------------------------------*

' Encodage du fichier texte
Private psEncodage As String
Private plEncodage As ListeEncodage
Private DicEncodage As Dictionary
' Separateur de lignes (fichier texte uniquement)
Private plSeparateurLigne As SeparateurLigne
' Type de fichier : Texte ou Binaire
Private plTypeFichier As TypeFichier
' Nom physique du fichier (le repertoire peut etre inclus dans le nom)
Private psNomFichier As String
' Objet ADODB.stream
Private poFichier As Object
' Type d'acces au fichier : lecture, ecriture
Private psTypeAcces As TypeAcces

' *---------------------------------------------------------------------------------------------------*
' * Gestion des boites de dialogue                                                                    *
' *---------------------------------------------------------------------------------------------------*

' Titre de la boite de dialogue (fichiers)
Private psTitreBoiteDialogue As String
' Libelle du filtre sur les fichiers a appliquer
Private psLibelleFiltre As String
' Extensions appliquees dans le filtre
Private psExtensionFiltre As String
' Nom du fichier initial propose dans la boite de dialogue Enregistrer sous
Private psNomInitialFichier As String
' Repertoire selectionne via la boite de dialogue
Private psRepertoire As String
' Filtre pour la boite a dialogue d'enregsitrement sous
Private plIndexFiltre As IndexFiltre

' *---------------------------------------------------------------------------------------------------*
' * Gestion des fichiers                                                                              *
' *---------------------------------------------------------------------------------------------------*

' Fichier ouvert
Private plStatutFlux As StatutFlux
Private plNbreEnregLus As Long
Private plNbreEnregEcrits As Long
Private plNbreOctectsLus As Long
Private plNbreOctectsEcrits As Long

' *---------------------------------------------------------------------------------------------------*
' * Propriete : objet fichier                                                                         *
' *---------------------------------------------------------------------------------------------------*
Property Set Fichier(oFichier As Object)
    Set poFichier = oFichier
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriete : type de fichier (texte, binaire)                                                      *
' *---------------------------------------------------------------------------------------------------*
Property Let TypeFichier(lTypeFichier As TypeFichier)

    ' Verifier le type de fichier
    Select Case lTypeFichier
        Case FICHIER_BINAIRE:
            Select Case True
                Case psEncodage <> "":
                    ' L'encodage n'est pas necessaire pour lire un fichier binaire
                    Err.Raise vbObjectError + 1000, , "[TypeFichier] L'encodage ne doit pas etre defini pour les fichiers binaires."
                Case plSeparateurLigne <> SEPARATEUR_NON_DEFINI:
                    ' Le separateur de lignes n'est pas necessaire pour lire un fichier binaire
                    Err.Raise vbObjectError + 1000, , "[TypeFichier] Le separateur de lignes ne doit pas etre defini pour les fichiers binaires."
                Case Else:
                    plTypeFichier = lTypeFichier
            End Select
        Case FICHIER_TEXTE:
            plTypeFichier = lTypeFichier
        Case Else
            Err.Raise vbObjectError + 1000, , "[TypeFichier] Le type de fichier " & lTypeFichier & " est invalide."
    End Select

End Property

Property Get TypeFichier() As TypeFichier
    TypeFichier = plTypeFichier
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriete : Encodage du fichier texte (Enum)                                                      *
' *---------------------------------------------------------------------------------------------------*
Property Let Encodage(lEncodage As ListeEncodage)

    Select Case True
        ' L'encodage doit etre renseigne
        Case lEncodage = NON_DEFINI:
            Err.Raise vbObjectError + 1000, , "[Encodage] L'encodage n'est pas renseigne."
        Case plTypeFichier = FICHIER_BINAIRE:
            ' L'encodage n'est pas necessaire pour lire un fichier binaire
            Err.Raise vbObjectError + 1000, , "[Encodage] L'encodage ne doit pas etre defini pour les fichiers binaires."
            ' Verifie que l'encodage existe dans l'enum
        Case EnumEncodageExiste(lEncodage) = True:
            plEncodage = lEncodage
            psEncodage = DicEncodage.Keys(lEncodage)
        Case Else:
            ' Si l'encodage est absent de l'enum ListeEncodage et du dictionnaire dicEncodage alors une erreur est retournee.
            ' La liste des encodages n'est pas exhaustive et doit etre completee si necessaire, sinon utiliser la propriete EncodageTxt.
            Err.Raise vbObjectError + 2, , "[Encodage] L'encodage n'existe pas dans l'enum ListeEncodage."
    End Select
    
End Property

Property Get Encodage() As ListeEncodage
    
    Encodage = plEncodage
    
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriete : Encodage du fichier texte (code technique)                                            *
' *---------------------------------------------------------------------------------------------------*
Property Let EncodageTxt(sEncodage As String)

    Select Case True
        ' L'encodage doit etre renseigne
        Case Trim(sEncodage) = "":
            Err.Raise vbObjectError + 3, , "[EncodageTxt] L'encodage n'est pas renseigne."
        ' L'encodage n'est pas necessaire pour lire un fichier binaire
        Case plTypeFichier = FICHIER_BINAIRE:
            Err.Raise vbObjectError + 1000, , "[EncodageTxt] L'encodage ne doit pas etre defini pour les fichiers binaires."
        ' Aucun controle de l'encodage, la saisie est libre.
        Case Else:
            psEncodage = sEncodage
            plEncodage = AUTRE
    End Select
    
End Property

Property Get EncodageTxt() As String
    EncodageTxt = psEncodage
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriete : separateur de lignes (CL, LF, CR/LF)                                                  *
' *---------------------------------------------------------------------------------------------------*
Property Let SeparateurLigne(lSeparateurLigne As SeparateurLigne)

    If lSeparateurLigne = SEPARATEUR_CR Or lSeparateurLigne = SEPARATEUR_CRLF Or lSeparateurLigne = SEPARATEUR_LF Then
        If plTypeFichier = FICHIER_BINAIRE Then
            ' Le separateur ne sert a rien pour les fichiers binaires
            Err.Raise vbObjectError + 1000, , "[SeparateurLigne] Le separateur ne doit pas etre defini pour les fichiers binaires."
        Else
            plSeparateurLigne = lSeparateurLigne
        End If
    Else
        Err.Raise vbObjectError + 1000, , "[SeparateurLigne] Le separateur " & lSeparateurLigne & " est invalide."
    End If
    
End Property

Property Get SeparateurLigne() As SeparateurLigne
    SeparateurLigne = plSeparateurLigne
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriete : type d'acces au fichier                                                               *
' *---------------------------------------------------------------------------------------------------*
Property Let TypeAcces(sTypeAcces As TypeAcces)

    If sTypeAcces = ACCES_ECRITURE Or sTypeAcces = ACCES_LECTURE Then
        psTypeAcces = sTypeAcces
    Else
        Err.Raise vbObjectError + 1000, , "[TypeAcces] Le type d'acces " & sTypeAcces & " est invalide."
    End If
    
End Property

Property Get TypeAcces() As TypeAcces
    TypeAcces = psTypeAcces
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriete : nom du fichier                                                                        *
' *---------------------------------------------------------------------------------------------------*
Property Let NomFichier(sNomFichier As String)

    If Trim(sNomFichier) = "" Then
        Err.Raise vbObjectError + 1000, , "[NomFichier] Le nom du fichier doit etre renseigne."
    Else
        ' Convertir les noms de fichier sous forme d'URL au format Lecteur:\Repertoire\NomFichier.extension
        psNomFichier = ConvertirUrlSharePoint(sNomFichier)
    End If
    
End Property

Property Get NomFichier() As String
    NomFichier = psNomFichier
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriete : titre de la boite de dialogue a afficher                                              *
' *---------------------------------------------------------------------------------------------------*
Property Let TitreBoiteDeDialogue(sTitre As String)

    If Trim(sTitre) = "" Then
        Err.Raise vbObjectError + 1000, , "[TitreBoiteDeDialogue] Le titre de la boite de dialogue doit etre renseigne."
    Else
        psTitreBoiteDialogue = sTitre
    End If
    
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriete : libelle du filtre a appliquer afin de selectionner les fichiers                       *
' *---------------------------------------------------------------------------------------------------*
Property Let LibelleFiltre(sLibelleFiltre As String)

    If Trim(sLibelleFiltre) = "" Then
        Err.Raise vbObjectError + 1000, , "[LibelleFiltre] Le libelle du filtre doit etre renseigne."
    Else
        psLibelleFiltre = sLibelleFiltre
    End If
    
End Property

Property Get LibelleFiltre() As String
    LibelleFiltre = psLibelleFiltre
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriete : Filtre des extensions separees par une virgule, eg. *.csv,*.txt                       *
' *---------------------------------------------------------------------------------------------------*
Property Let ExtensionFiltre(sExtensionFiltre As String)

    If Trim(sExtensionFiltre) = "" Then
        Err.Raise vbObjectError + 1000, , "[ExtensionFiltre] Les extensions doivent etre renseignees."
    Else
        psExtensionFiltre = sExtensionFiltre
    End If
    
End Property

Property Get ExtensionFiltre() As String
    ExtensionFiltre = psExtensionFiltre
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriete : nom du fichier pre renseigne dans la boite de dialogue pour enregistrer sous          *
' *---------------------------------------------------------------------------------------------------*
Property Let NomInitialFichier(sNomInitialFichier As String)

    If Trim(sNomInitialFichier) = "" Then
        Err.Raise vbObjectError + 1000, , "[NomInitialFichier] Le nom initial du fichier n'est pas renseigne."
    Else
        psNomInitialFichier = sNomInitialFichier
    End If

End Property

Property Get NomInitialFichier() As String
    NomInitialFichier = psNomInitialFichier
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriete : chemin d'un repertoire (dossier)                                                      *
' *---------------------------------------------------------------------------------------------------*
Property Let Repertoire(sRepertoire As String)
    
    If Trim(sRepertoire) = "" Then
        Err.Raise vbObjectError + 1000, , "[Repertoire] Le repertoire n'est pas renseigne."
    End If
    
    psRepertoire = sRepertoire
    
End Property

Property Get Repertoire() As String
    Repertoire = psRepertoire
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriete : Filtre a appliquer lors de la selection des fichiers a enregistrer sous               *
' *---------------------------------------------------------------------------------------------------*
Property Let Filtre(lIndexFiltre As IndexFiltre)
    
    If lIndexFiltre < 1 Or lIndexFiltre > 29 Then
        Err.Raise vbObjectError + 1000, , "[lIndexFiltre] Le filtre " & lIndexFiltre & " est invalide."
    Else
        plIndexFiltre = lIndexFiltre
    End If
    
End Property

Property Get Filtre() As IndexFiltre
    Filtre = plIndexFiltre
End Property

' *---------------------------------------------------------------------------------------------------*
' * Ouverture d'un fichier via ADODB                                                                  *
' *---------------------------------------------------------------------------------------------------*
' * Proprietes utilisees :                                                                            *
' *   -- TypeFichier     (obligatoire)                                                                *
' *   -- SeparateurLigne (obligatoire si TypeFichier = Texte et lecture par ligne)                    *
' *   -- Encodage        (obligatoire si TypeFichier = Texte)                                         *
' *   -- TypeAcces       (obligatoire)                                                                *
' *   -- NomFichier      (obligatoire si acces en lecture)                                            *
' *---------------------------------------------------------------------------------------------------*
Public Sub Ouvrir()
    
    Select Case True
        ' Verifier le type de fichier
        Case plTypeFichier = FICHIER_NON_DEFINI:
            Err.Raise vbObjectError + 1000, , "[Ouvrir] Le type de fichier n'est pas defini."
        ' Verifier que l'encodage est renseigne si le fichier est de type Texte
        plTypeFichier = FICHIER_TEXTE And psEncodage = "":
            Err.Raise vbObjectError + 1000, , "[Ouvrir] L'encodage n'est pas defini."
        ' Verifier que le type d'acces est renseigne
        Case psTypeAcces = ACCES_NON_DEFINI:
            Err.Raise vbObjectError + 1000, , "[Ouvrir] Le type d'acces n'est pas defini."
        Case psTypeAcces = ACCES_LECTURE:
            ' Si le nom de fichier n'est pas renseigne
            If Trim(psNomFichier) = "" Then
                Err.Raise vbObjectError + 1000, , "[Ouvrir] Le nom du fichier n'est pas defini."
            End If
            ' Verifier que le fichier existe si l'acces est en lecture
            If FichierExiste = False Then
                Err.Raise vbObjectError + 1000, , "[Ouvrir] Le fichier " & psNomFichier & " n'existe pas."
            End If
    End Select
    
    With poFichier
        ' Type de fichier : binaire ou texte
        .Type = plTypeFichier
        ' Si fichier texte, preciser l'encodage et le separateur
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
' * Lecture des n prochains caracteres d'un fichier texte ADODB                                       *
' *---------------------------------------------------------------------------------------------------*
' * Parametre utilise :                                                                               *
' *   -- lNbreCarALire    : nombre de caracteres a lire dans le fichier ou jusqu'a la fin de fichier  *
' *---------------------------------------------------------------------------------------------------*
Public Function Lire(lNbreCarALire As Long) As String
    
    Select Case True
        ' Acces impossible a un fichier non ouvert
        Case plStatutFlux <> STATUT_FICHIER_OUVERT:
            Err.Raise vbObjectError + 1000, , "[Lire] Le fichier n'a pas ete ouvert."
        ' Le nombre de caracteres a lire doit etre > 0
        Case lNbreCarALire <= 0:
            Err.Raise vbObjectError + 1000, , "[Lire] Le nombre de caracteres a lire n'est pas valide."
        ' Seule la lecture complete d'un fichier binaire est possible
        Case plTypeFichier = FICHIER_BINAIRE:
            Err.Raise vbObjectError + 1000, , "[Lire] La lecture partielle d'un fichier binaire n'est pas possible."
        Case Else:
            ' Lire les n caracteres suivants
            Lire = poFichier.ReadText(lNbreCarALire)
            plNbreOctectsLus = plNbreOctectsLus + Len(Lire)
    End Select
    
End Function

' *---------------------------------------------------------------------------------------------------*
' * Lecture du prochain enregistrement d'un fichier texte ADODB                                       *
' *---------------------------------------------------------------------------------------------------*
' * Propriete utilisee :                                                                              *
' *   -- SeparateurLigne                                                                              *
' *---------------------------------------------------------------------------------------------------*
Public Function LireEnregistrement() As String
    
    Select Case True
        ' Acces impossible a un fichier non ouvert
        Case plStatutFlux <> STATUT_FICHIER_OUVERT:
            Err.Raise vbObjectError + 1000, , "[LireEnregistrement] Le fichier n'a pas ete ouvert."
        ' Verifie que le separateur a bien ete defini pour la lecture par ligne
        Case plSeparateurLigne = SEPARATEUR_NON_DEFINI:
            Err.Raise vbObjectError + 1000, , "[LireEnregistrement] Le separateur de ligne n'est pas defini."
        ' Lecture du prochain enregistrement
        Case plTypeFichier = FICHIER_TEXTE:
            LireEnregistrement = poFichier.ReadText(LIRE_ENREG_SUIVANT)
            plNbreEnregLus = plNbreEnregLus + 1
            plNbreOctectsLus = plNbreOctectsLus + Len(LireEnregistrement)
        ' Les fichiers binaires ne comportent pas de separateur, par consequent, il n'est pas possible de lire un enregistrement
        Case plTypeFichier = FICHIER_BINAIRE:
            Err.Raise vbObjectError + 1000, , "[LireEnregistrement] La lecture d'un enregistrement d'un fichier binaire n'est pas possible."
    End Select
    
End Function

' *---------------------------------------------------------------------------------------------------*
' * Fin de fichier atteinte ?                                                                         *
' *---------------------------------------------------------------------------------------------------*
Public Function FinFichier() As Boolean
    
    ' Acces impossible a un fichier non ouvert
    If plStatutFlux <> STATUT_FICHIER_OUVERT Then
        Err.Raise vbObjectError + 1000, , "[FinFichier] Le fichier n'a pas ete ouvert."
    End If
    
    FinFichier = poFichier.EOS

End Function

' *---------------------------------------------------------------------------------------------------*
' * Lecture integrale d'un fichier ADODB                                                              *
' *---------------------------------------------------------------------------------------------------*
Public Function LireFichier() As Variant
    
    Select Case True
        ' Acces impossible a un fichier non ouvert
        Case plStatutFlux <> STATUT_FICHIER_OUVERT:
            Err.Raise vbObjectError + 1000, , "[LireFichier] Le fichier n'a pas ete ouvert."
        ' Lire entierement un fichier binaire
        Case plTypeFichier = FICHIER_BINAIRE:
            LireFichier = poFichier.Read(LIRE_FICHIER)
            plNbreOctectsLus = plNbreOctectsLus + Len(LireFichier)
        ' Lire entierement un fichier texte
        Case plTypeFichier = FICHIER_TEXTE:
            LireFichier = poFichier.ReadText(LIRE_FICHIER)
            plNbreOctectsLus = plNbreOctectsLus + Len(LireFichier)
    End Select

End Function

' *---------------------------------------------------------------------------------------------------*
' * Ecrire un enregistrement dans un flux ADODB.stream                                                *
' *---------------------------------------------------------------------------------------------------*
' * Parametres :                                                                                      *
' *   -- sEnregistrement (Obligatoire) : enregistrement a ecrire (XML, json, texte, etc)              *
' *---------------------------------------------------------------------------------------------------*
Public Sub EcrireEnregistrement(sEnregistrement As String)
    
    Select Case True
        ' Acces impossible a un fichier non ouvert
        Case plStatutFlux <> STATUT_FICHIER_OUVERT:
            Err.Raise vbObjectError + 26, , "[EcrireEnregistrement] Le fichier n'a pas ete ouvert."
        ' Ecriture de l'enregistrement dans le fichier texte. Les separateurs CR/LF doivent etre presents dans l'enregistrement.
        Case plTypeFichier = FICHIER_TEXTE:
            poFichier.WriteText sEnregistrement
            plNbreEnregEcrits = plNbreEnregEcrits + 1
            plNbreOctectsEcrits = plNbreOctectsEcrits + Len(sEnregistrement)
        ' Il n'est pas possible d'ecrire un enregistrement dans un fichier binaire.
        Case plTypeFichier = FICHIER_BINAIRE:
            Err.Raise vbObjectError + 1000, , "[EcrireEnregistrement] L'ecriture d'un enregistrement dans un fichier binaire n'est pas possible."
    End Select
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Ecrire des donnees dans un flux ADODB.stream                                                      *
' *---------------------------------------------------------------------------------------------------*
' * Parametres :                                                                                      *
' *   -- vDonnees (Obligatoire) : donnees a ecrire (image, pdf, etc)                                  *
' *---------------------------------------------------------------------------------------------------*
Public Sub Ecrire(vDonnees As Variant)
    
    Select Case True
        ' Acces impossible a un fichier non ouvert
        Case plStatutFlux <> STATUT_FICHIER_OUVERT:
            Err.Raise vbObjectError + 26, , "[Ecrire] Le fichier n'a pas ete ouvert."
        ' Ecrire le fichier binaire (en une seule fois)
        Case plTypeFichier = FICHIER_BINAIRE:
            poFichier.Write vDonnees
            plNbreOctectsEcrits = plNbreOctectsEcrits + Len(vDonnees)
        ' Il n'est pas possible d'ecrire des donnees brutes dans un fichier texte.
        Case plTypeFichier = FICHIER_TEXTE:
            Err.Raise vbObjectError + 1000, , "[Ecrire] L'ecriture de donnees binaires dans fichier texte n'est pas possible."
    End Select
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Enregistrer le flux ADODB.Stream sous le nom du fichier                                           *
' *---------------------------------------------------------------------------------------------------*
' * Proprietes utilisees :                                                                            *
' *   -- TypeFichier     (obligatoire)                                                                *
' *   -- NomFichier      (obligatoire)                                                                *
' *---------------------------------------------------------------------------------------------------*
Public Sub EnregistrerSous()

    Select Case True
        Case plStatutFlux <> STATUT_FICHIER_OUVERT:
            ' Acces impossible a un fichier non ouvert
            Err.Raise vbObjectError + 26, , "[EnregistrerSous] Le fichier n'a pas ete ouvert."
        Case Trim(psNomFichier) = "":
            Err.Raise vbObjectError + 1000, , "[EnregistrerSous] Le nom du fichier n'est pas defini."
        Case plTypeFichier = FICHIER_NON_DEFINI:
            Err.Raise vbObjectError + 1000, , "[EnregistrerSous] Le type de fichier n'est pas defini."
        Case Else
            poFichier.SaveToFile psNomFichier, plTypeFichier
    End Select
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Fermer un objet ADODB                                                                             *
' *---------------------------------------------------------------------------------------------------*
Public Sub Fermer()
    
    If plStatutFlux <> STATUT_FICHIER_OUVERT Then
        ' Acces impossible a un fichier non ouvert
        Err.Raise vbObjectError + 26, , "[Fermer] Le fichier n'a pas ete ouvert."
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
' * Affiche la boite de dialogue qui permet de selectionner le fichier afin d'enregistrer sous        *
' *---------------------------------------------------------------------------------------------------*
' * Proprietes utilisees :                                                                            *
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
        ' Titre par defaut de la boite de dialogue
        If psTitreBoiteDialogue = "" Then
            .Title = "Selectionner le repertoire et saisir le nom du fichier a enregistrer sous"
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
' * Affiche la boite de dialogue qui permet de selectionner un fichier qui correspond aux filtres     *
' *---------------------------------------------------------------------------------------------------*
' * Proprietes utilisees :                                                                            *
' *   -- LibelleFiltre       (Facultatif)                                                             *
' *   -- ExtensionFiltre     (Facultatif)                                                             *
' *   -- TitreBoiteDialogue  (Facultatif)                                                             *
' *---------------------------------------------------------------------------------------------------*
Public Sub SelectionnerFichier()

    ' Boite de dialogue de selection d'un fichier
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        If psNomInitialFichier <> "" Then
            .InitialFileName = psNomInitialFichier
        End If
        .InitialView = msoFileDialogViewDetails
        ' Si filtre ou extension non renseigne
        If psLibelleFiltre = "" Or psExtensionFiltre = "" Then
            .Filters.Add "Tous les fichiers", "*.*", 1
        Else
            .Filters.Add psLibelleFiltre, psExtensionFiltre, 1
        End If
        .FilterIndex = 1
        ' Titre par defaut si non renseigne
        If psTitreBoiteDialogue = "" Then
            .Title = "Selectionner le repertoire et le fichier"
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
' Affiche la boite de dialogue qui permet de selectionner un repertoire                               *
' *---------------------------------------------------------------------------------------------------*
' * Propriete utilisee :                                                                              *
' *   -- TitreBoiteDialogue  (Facultatif)                                                             *
' *---------------------------------------------------------------------------------------------------*
Public Sub SelectionnerRepertoire()

    ' Boite de dialogue de selection du repertoire
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        If psNomInitialFichier <> "" Then
            .InitialFileName = psNomInitialFichier
        End If
        .InitialView = msoFileDialogViewDetails
        If psTitreBoiteDialogue = "" Then
            .Title = "Selectionner le repertoire"
        Else
            .Title = psTitreBoiteDialogue
        End If
        ' Si aucun repertoire selectionne alors fin du traitement
        If .Show = -1 Then
            psRepertoire = .SelectedItems(1)
        Else
            psRepertoire = ""
        End If
    End With

End Sub

' *---------------------------------------------------------------------------------------------------*
' * Verifie si un fichier est trop volumineux pour etre traite par ADODB                              *
' *---------------------------------------------------------------------------------------------------*
' * Propriete utilisee :                                                                              *
' *   -- NomFichier          (Obligatoire)                                                            *
' *---------------------------------------------------------------------------------------------------*
Public Function EstFichierVolumineux(Optional sNomFichier As String) As Boolean
    
    Dim sNomFichierLocal As String
    
    Select Case True
        Case Trim(sNomFichier) <> "": sNomFichierLocal = sNomFichier
        Case psNomFichier <> "": sNomFichierLocal = psNomFichier
        Case Else: Err.Raise vbObjectError + 1000, , "EstFichierVolumineux : Le nom du fichier n'est pas defini."
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
' * Parametre utilise :                                                                               *
' *   -- sNomFichier         (Facultatif)                                                             *
' * Propriete utilisee :                                                                              *
' *   -- NomFichier          (Facultatif)                                                             *
' *---------------------------------------------------------------------------------------------------*
Public Function LongueurFichier(Optional sNomFichier As String) As Double

    Dim sNomFichierLocal As String
    Dim oFileSystem As Object, oFichier As Object
    
    Select Case True
        Case Trim(sNomFichier) <> "": sNomFichierLocal = sNomFichier
        Case psNomFichier <> "": sNomFichierLocal = psNomFichier
        Case Else: Err.Raise vbObjectError + 1000, , "LongueurFichier : Le nom du fichier n'est pas defini."
    End Select

    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
    Set oFichier = oFileSystem.GetFile(sNomFichierLocal)
    LongueurFichier = oFichier.Size

End Function

' *---------------------------------------------------------------------------------------------------*
' * Verifie si un fichier existe                                                                      *
' *---------------------------------------------------------------------------------------------------*
' * Parametre utilise :                                                                               *
' *   -- sNomFichier         (Facultatif)                                                             *
' * Propriete utilisee :                                                                              *
' *   -- NomFichier          (Facultatif)                                                             *
' *---------------------------------------------------------------------------------------------------*
Public Function FichierExiste(Optional sNomFichier As String) As Boolean
    
    Dim sNomFichierLocal As String
    
    Select Case True
        Case Trim(sNomFichier) <> "": sNomFichierLocal = sNomFichier
        Case psNomFichier <> "": sNomFichierLocal = psNomFichier
        Case Else: Err.Raise vbObjectError + 1000, , "FichierExiste : le nom du fichier n'est pas defini."
    End Select

    FichierExiste = Dir(sNomFichierLocal, vbNormal) <> ""

End Function

' *---------------------------------------------------------------------------------------------------*
' * Verifie si un repertoire existe                                                                   *
' *---------------------------------------------------------------------------------------------------*
' * Propriete utilisee :                                                                              *
' *   -- Repertoire          (Obligatoire)                                                            *
' *---------------------------------------------------------------------------------------------------*
Public Function RepertoireExiste(Optional sRepertoire As String) As Boolean
    
    Dim sNomRepertoire As String
    
    Select Case True
        Case Trim(sRepertoire) <> "": sNomRepertoire = sRepertoire
        Case psRepertoire <> "": sNomRepertoire = psRepertoire
        Case Else: Err.Raise vbObjectError + 1000, , "RepertoireExiste : le nom du repertoire n'est pas defini."
    End Select

    RepertoireExiste = Dir(sNomRepertoire, vbDirectory) <> ""

End Function

' *---------------------------------------------------------------------------------------------------*
' * Verifie si l'encodage existe dans le dictionnaire                                                 *
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
' * Convertir un nom de chemin defini par une URL OneDrive ou SharePoint vers un nom de chemin Windows                       *
' * Exemple : https://xxx-my.sharepoint.com/personal/ devient c:\Users\xxxx\OneDrive - xxx                                   *
' *--------------------------------------------------------------------------------------------------------------------------*
Private Function ConvertirUrlSharePoint(sChemin As String) As String

    Dim sListeDossiers() As String, iNbDossiers As Integer, lPosDoc As Long, sRepertoire As String
    
    ' Si le chemin du fichier commence par http
    If LCase$(Left(sChemin, 4)) = "http" Then
        Select Case True
        ' Espace personnel sur SharePoint (i.e. OneDrive Commercial) ?
        Case sChemin Like "https://*-my.sharepoint.com/personal/":
            ' Recherche la chaine "/Documents/documents" afin d'obtenir le debut de l'arborescence dans le dossier des documents
            lPosDoc = InStr(1, sChemin, "/Documents/Documents/", vbTextCompare) + Len("/Documents")
            ' Le repertoire local est recupere a partir du 2eme /Documents
            sRepertoire = Mid(sChemin, lPosDoc, Len(sChemin) - lPosDoc + 1)
            ConvertirUrlSharePoint = Environ("OneDriveCommercial") & Replace(sRepertoire, "/", "\")
        ' Espace de travail partage
        Case sChemin Like "https://weshare*":
            sListeDossiers = Split(sChemin, "/")
            ConvertirUrlSharePoint = "\\" & sListeDossiers(2) & "@SSL\DavWWWRoot"
            For iNbDossiers = 3 To UBound(sListeDossiers)
                ConvertirUrlSharePoint = ConvertirUrlSharePoint & "\" & sListeDossiers(iNbDossiers)
            Next
        Case sChemin Like "https://d.docs.live.net/*":
            ' Recherche la chaine "/documents" afin d'obtenir le debut de l'arborescence dans le dossier des documents
            lPosDoc = InStr(1, sChemin, "/Documents/", vbTextCompare)
            ' Le repertoire local est recupere a partir du 2eme /Documents
            sRepertoire = Mid(sChemin, lPosDoc, Len(sChemin) - lPosDoc + 1)
            ConvertirUrlSharePoint = Environ("OneDrive") & Replace(sRepertoire, "/", "\")
        End Select
    Else
        ConvertirUrlSharePoint = sChemin
    End If
    
End Function

' *---------------------------------------------------------------------------------------------------*
' * Initialisation de la classe : creation des objets ADODB.Stream et le dictionnaire des encodages   *
' *---------------------------------------------------------------------------------------------------*
Private Sub Class_Initialize()
    
    ' Creation d'un objet ADODB
    Set poFichier = CreateObject("ADODB.STREAM")

    ' Creation d'un dictionnaire des encodages
    Set DicEncodage = New Dictionary
    ' Ajout des encodages (a completer si necessaire)
    ' cle : libelle de l'encodage identique aux cles de registre charset de windows
    ' La liste ci-dessous doit etre identique a l'enum ListeEncodage
    DicEncodage.Add "Non defini", 0       ' NON_DEFINI
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
' * Terminaison de la classe : liberer la memoire en detruisant les objets                            *
' *---------------------------------------------------------------------------------------------------*
Private Sub Class_Terminate()
    Set poFichier = Nothing
    Set DicEncodage = Nothing
End Sub
