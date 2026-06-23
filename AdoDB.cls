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
' Module de classe : AdoDB
' Descritpion      : Encapsulage des fonctionnalités AdoDB.Stream dans une classe VBA afin de gérer des fichiers texte et binaire.
'                    La classe permet de faire abstraction de ADODB en proposant des domaines de valeurs pour chaque propriété.
'                    Les contrôles réalisés pour chaque propriété et méthode sécurise les appels à AdoDB en signalant les erreurs
'                    en amont.
' Limites connues  : La gestion des fichiers de plus de 4 Go n'est pas gérée correctement par AdoDB. Lors de la lecture de gros
'                    fichiers, une perte de données est généralement constatée. Cela se manifeste par des enregistrements qui
'                    sont ignorés et par conséquent absents lors de la lecture de tous les enregistrements d'un fichier texte.
'-------------------------------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------------------------------
' Propriétés :
' +--------------------------------+--------------------+----------------------------------------------------------+
' | Nom                            | Type               | Description                                              |
' +--------------------------------+--------------------+----------------------------------------------------------+
' | TypeFichier                    | StreamTypeEnum     | Type du fichier (texte ou binaire)                       |
' | Encodage                       | CharsetEnum        | Encodage du fichier                                      |
' | EncodageTxt                    | String             | Encodage libre (UTF-8, ASCII…)                           |
' | SeparateurLigne                | LineSeparatorsEnum | Séparateur de lignes d'un fichier texte                  |
' | ModeAcces                      | ConnectModeEnum    | Accès au fichier (lecture, écriture, MAJ)                |
' | NomFichier                     | String             | Chemin complet du fichier                                |
' | TitreBoiteDeDialogue           | String             | Titre des boîtes de dialogue                             |
' | LibelleFiltre                  | String             | Libellé du filtre                                        |
' | ExtensionFiltre                | String             | Filtre des extensions de fichier utilisé dans une boîte  |
' |                                |                    | de dialogue                                              |
' | NomInitialFichier              | String             | Nom proposé dans Enregistrer sous                        |
' | Repertoire                     | String             | Répertoire sélectionné                                   |
' | Filtre                         | FilterIndexEnum    | Filtre Excel (xlsx, csv, txt…)                           |
' | Fichier                        | ADODB.Stream       | Permet d’utiliser un Stream externe                      |
' | NbreEnregLus                   | Long               | Nombre d’enregistrements lus                             |
' | NbreEnregEcrits                | Long               | Nombre d’enregistrements écrits                          |
' | NbreOctectsLus                 | Long               | Octets lus                                               |
' | NbreOctectsEcrits              | Long               | Octets écrits                                            |
' +--------------------------------+--------------------+----------------------------------------------------------+
'
' Méthodes :
' +--------------------------------+-------------------------------------------------------------------------------+
' | Nom                            | Description                                                                   |
' +--------------------------------+-------------------------------------------------------------------------------+
' | Ouvrir                         | Ouvre un fichier via ADODB.Stream                                             |
' | Lire                           | Lit n caractères                                                              |
' | LireEnregistrement             | Lit jusqu’au prochain séparateur                                              |
' | LireFichier                    | Lit tout le fichier                                                           |
' | FinFichier                     | Indique si la fin est atteinte                                                |
' | Position                       | Déplace le pointeur                                                           |
' | Ecrire                         | Écrit des données                                                             |
' | EcrireEnregistrement           | Écrit un enregistrement texte                                                 |
' | EnregistrerSous                | Sauvegarde le flux dans un fichier                                            |
' | Fermer                         | Ferme le flux                                                                 |
' | SelectionnerFichier            | Ouvre une boîte de dialogue                                                   |
' | SelectionnerRepertoire         | Sélectionne un dossier                                                        |
' | EstFichierVolumineux           | Vérifie > 4 Go                                                                |
' | FichierExiste                  | Vérifie l’existence d'un fichier                                              |
' | RepertoireExiste               | Vérifie l’existence d'un répertoire                                           |
' | LongueurFichier                | Retourne la taille d'un fichier                                               |
' | Extension                      | Retourner l'extension du fichier (en minuscules)                              |
' +--------------------------------+-------------------------------------------------------------------------------+
'

Option Explicit
Option Compare Text

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
' Propriété : Fichier
' Rôle      : Pointer sur un objet AdoDB externe à la classe
' Paramètre : FichierAdoDB
' Erreurs   :
'-------------------------------------------------------------------------------------------------------------------------
Property Set Fichier(FichierAdodb As Object)

    Set pFichierADODB = FichierAdodb

End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : TypeFichier
' Rôle      : Définit le type de fichier (texte ou binaire).
' Paramètre : TypeFichier (StreamTypeEnum)
' Erreurs   : -100 encodage défini pour un fichier binaire
'             -101 séparateur défini pour un fichier binaire
'             -102 type de fichier invalide
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
' Propriété : Encodage
' Rôle      : Définit l'encodage du fichier texte via CharsetEnum.
' Paramètre : EncodageNum (CharsetEnum)
' Erreurs   : -110 encodage non renseigné
'             -111 encodage défini pour un fichier binaire
'             -112 encodage absent de l'enum
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
' Propriété : EncodageTxt
' Rôle      : Définit l'encodage libre (UTF-8, ASCII…).
'             Aucun contrôle n'est réalisé, l'encodage est transmis tel quel à AdoDB qui va retourner une erreur s'il n'est pas valide.
' Paramètre : Encodage (String)
' Erreurs   : -120 encodage vide
'             -121 encodage défini pour un fichier binaire
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
' Propriété : SeparateurLigne
' Rôle      : Définit le séparateur de lignes pour les fichiers texte.
' Paramètre : SeparateurLigne (LineSeparatorsEnum)
' Erreurs   : -130 séparateur défini pour un fichier binaire
'             -131 séparateur invalide
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
' Propriété : ModeAcces
' Rôle      : Définit le mode d'accès au fichier (lecture/écriture).
' Paramètre : ModeAcces (ConnectModeEnum)
' Erreurs   : -140 mode invalide
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
' Propriété : NomFichier
' Rôle      : Définit le chemin complet du fichier à traiter.
' Paramètre : NomFichier (String)
' Erreurs   : -150 nom vide
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
' Propriété : TitreBoiteDeDialogue
' Rôle      : Titre de la boîte de dialogue qui permet de sélectionner un fichier / répertoire
' Paramètre : Titre (string)
' Erreurs   : -160 Titre non renseigné
'-------------------------------------------------------------------------------------------------------------------------
Property Let TitreBoiteDeDialogue(Titre As String)

    If Trim(Titre) = "" Then
        Err.Raise -160, "AdoDB.TitreBoiteDeDialogue", "Le titre de la boîte de dialogue doit être renseigné."
    Else
        pTitreBoiteDialogue = Titre
    End If
    
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : LibelleFiltre
' Rôle      : Libellé du filtre à appliquer afin de sélectionner les fichiers
' Paramètre : LibelleFiltre (string)
' Erreurs   : -170 Libellé non renseigné
'-------------------------------------------------------------------------------------------------------------------------
Property Let LibelleFiltre(LibelleFiltre As String)

    If Trim(LibelleFiltre) = "" Then
        Err.Raise -170, "AdoDB.LibelleFiltre", "Le libellé du filtre doit être renseigné."
    Else
        pLibelleFiltre = LibelleFiltre
    End If
    
End Property

Property Get LibelleFiltre() As String
    LibelleFiltre = pLibelleFiltre
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : ExtensionFiltre
' Rôle      : Filtre des extensions séparées par une virgule, eg. *.csv,*.txt
' Paramètre : ExtensionFiltre (String)
' Erreurs   : -180 Extension non renseignée
'-------------------------------------------------------------------------------------------------------------------------
Property Let ExtensionFiltre(ExtensionFiltre As String)

    If Trim(ExtensionFiltre) = "" Then
        Err.Raise -180, "AdoDB.ExtensionFiltre", "Les extensions doivent être renseignées."
    Else
        pExtensionFiltre = ExtensionFiltre
    End If
    
End Property

Property Get ExtensionFiltre() As String
    ExtensionFiltre = pExtensionFiltre
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : NomInitialFichier
' Rôle      : Nom du fichier prérenseigné dans la boite de dialogue pour enregistrer sous
' Paramètre : NomInitialFichier (String)
' Erreurs   : -190 Nom initial non renseigné
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
' Propriété : Repertoire
' Rôle      : Définit un répertoire ou récupère celui sélectionné dans la boite de dialogue
' Paramètre : Repertoire (string)
' Erreurs   : -200 Répertoire non renseigné
'-------------------------------------------------------------------------------------------------------------------------
Property Let Repertoire(Repertoire As String)
    
    If Trim(Repertoire) = "" Then
        Err.Raise -200, "AdoDB.Repertoire", "Le répertoire n'est pas renseigné."
    Else
        pRepertoire = Repertoire
    End If
    
End Property

Property Get Repertoire() As String
    Repertoire = pRepertoire
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : Filtre
' Rôle      : Filtre à appliquer lors de la sélection des fichiers à enregistrer sous
' Paramètre : IndexFiltre (FilterIndexEnum)
' Erreurs   : -210 Index invalide
'-------------------------------------------------------------------------------------------------------------------------
Property Let Filtre(IndexFiltre As FilterIndexEnum)

    If IndexFiltre < 0 Or IndexFiltre > 29 Then
        Err.Raise -210, "AdoDB.Filtre", "Le filtre " & IndexFiltre & " est invalide."
    Else
        pIndexFiltre = IndexFiltre
    End If
    
End Property

Property Get Filtre() As FilterIndexEnum
    Filtre = pIndexFiltre
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : NbreEnregLus
' Rôle      : Retourne le nombre d'enregistrements lus (si fichier texte)
' Paramètre : N/A
' Erreurs   : N/A
'-------------------------------------------------------------------------------------------------------------------------
Property Get NbreEnregLus() As Long
    NbreEnregLus = pNbreEnregLus
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : NbreEnregEcrits
' Rôle      : Retourne le nombre d'enregistrements écrits (si fichier texte)
' Paramètre : N/A
' Erreurs   : N/A
'-------------------------------------------------------------------------------------------------------------------------
Property Get NbreEnregEcrits() As Long
    NbreEnregEcrits = pNbreEnregEcrits
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : NbreOctectsLus
' Rôle      : Retourne le nombre d'octets lus (si fichier texte)
' Paramètre : N/A
' Erreurs   : N/A
'-------------------------------------------------------------------------------------------------------------------------
Property Get NbreOctectsLus() As Long
    NbreOctectsLus = pNbreOctectsLus
End Property

'-------------------------------------------------------------------------------------------------------------------------
' Propriété : NbreOctectsEcrits
' Rôle      : Retourne le nombre d'octets écrits (si fichier texte)
' Paramètre : N/A
' Erreurs   : N/A
'-------------------------------------------------------------------------------------------------------------------------
Property Get NbreOctectsEcrits() As Long
    NbreOctectsEcrits = pNbreOctectsEcrits
End Property

'------------------------------------------------------------------------------
' Méthode   : Ouvrir
' Rôle      : Ouvre un fichier via ADODB.Stream selon les propriétés définies.
' Paramètre : N/A
' Propriétés obligatoires :
'   - TypeFichier
'   - Encodage (si texte)
'   - SeparateurLigne (si lecture par ligne)
'   - ModeAcces
'   - NomFichier (si lecture)
' Erreurs   : -220 type de fichier non défini
'             -221 encodage non défini
'             -222 mode d'accès non défini
'             -223 nom de fichier non défini
'             -224 fichier inexistant
'------------------------------------------------------------------------------
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
' Méthode   : Lire
' Rôle      : Lit les n caractères suivants du fichier
' Propriétés obligatoires :
'   - TypeFichier
' Paramètre : NbreCarALire (long)
' Résultat  : n caractères lus
' Erreurs   : -230 fichier non ouvert
'             -231 nombre de caractères invalide
'             -232 lecture partielle impossible en binaire
'-------------------------------------------------------------------------------------------------------------------------
Public Function Lire(NbreCarALire As Long) As String
    
    Select Case True
        ' Accès impossible à un fichier non ouvert
        Case pStatutFichier <> OPENED_FILE
            Err.Raise -230, "AdoDB.Lire", "Le fichier n'a pas été ouvert."
        ' Le nombre de caractères à lire doit être > 0
        Case NbreCarALire <= 0
            Err.Raise -231, "AdoDB.Lire", "Le nombre de caractères à lire n'est pas valide."
        ' Seule la lecture complète d'un fichier binaire est possible
        Case pTypeFichier = AD_TYPE_BINARY
            Err.Raise -232, "AdoDB.Lire", "La lecture partielle d'un fichier binaire n'est pas possible."
        Case Else
            ' Lire les n caractères suivants
            Lire = pFichierADODB.ReadText(NbreCarALire)
            pNbreOctectsLus = pNbreOctectsLus + Len(Lire)
    End Select
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Méthode   : LireEnregistrement
' Rôle      : Lit un enregistrement jusqu'au prochain séparateur.
' Propriétés obligatoires :
'   - SeparateurLigne
'   - TypeFichier
' Erreurs   : -240 fichier non ouvert
'             -241 séparateur non défini
'             -242 lecture impossible en binaire
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
' Méthode   : FinFichier
' Rôle      : fin du fichier atteinte
' Paramètre : N/A
' Résultat  : Vrai si la fin du fichier est atteinte sinon Faux
' Erreurs   : -250 Fichier non ouvert
'-------------------------------------------------------------------------------------------------------------------------
Public Function FinFichier() As Boolean
    
    ' Accès impossible à un fichier non ouvert
    If pStatutFichier <> OPENED_FILE Then
        Err.Raise -250, "AdoDB.FinFichier", "Le fichier n'a pas été ouvert."
    End If
    
    FinFichier = pFichierADODB.EOS

End Function

'-------------------------------------------------------------------------------------------------------------------------
' Méthode   : Position
' Rôle      : Repositionne sur le n-ième caractère
' Paramètre : Position (long)
' Résultat  : La prochaine lecture / écriture sera réalisée à la position inidiquée
' Erreurs   : N/A
'-------------------------------------------------------------------------------------------------------------------------
Public Sub Position(Position As Long)

    pFichierADODB.Position = Position
    
End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Méthode   : LireFichier
' Rôle      : Lit l'intégralité du fichier (texte ou binaire).
' Propriétés obligatoires :
'   - TypeFichier
' Paramètre : N/A
' Erreurs   : -260 fichier non ouvert
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
' Méthode   : EcrireEnregistrement
' Rôle      : Ecriture d'un enregistrement (fichier texte) dans le flux AdoDB.Stream
' Paramètre : Enregistrement (string)
' Propriétés obligatoires :
'   - TypeFichier
' Résultat  : Ecriture du flux AdoDB.Stream
' Erreurs   : -270 Fichier non ouvert
'             -271 Fichier binaire invalide
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
' Méthode   : Ecrire
' Rôle      : Ecriture de données (fichier binaire) dans le flux AdoDB.Stream
' Paramètre : Donnees (variant)
' Propriétés obligatoires :
'   - TypeFichier
' Résultat  : Ecriture du flux AdoDB.Stream
' Erreurs   : -280 fichier non ouvert
'             -281 Fichier texte invalide
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
' Méthode   : EnregistrerSous
' Rôle      : Enregistrer le flux AdoDB.stream dans un fichier physique avec écrasement
' Paramètre : N/A
' Propriétés obligatoires :
'   - NomFichier
'   - TypeFichier
'   - Encodage
' Résultat  : Fichier physique écrit
' Erreurs   : -290 fichier non ouvert
'             -291 Nom de fichier non renseigné
'             -292 type de fichier non défini
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
' Méthode   : fermer
' Rôle      : Fermer le flux AdoDB.Stream
' Paramètre : N/A
' Résultat  : Flux fermé
' Erreurs   : -300 fichier non ouvert
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
' Méthode   : SelectionnerFichierEnregistrerSous
' Rôle      : Affiche la boite de dialogue qui permet de sélectionner le fichier afin d'enregistrer sous
' Paramètre : N/A
' Résultat  : renseigne le nom du fichier
' Erreurs   : N/A
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
' Méthode   : SelectionnerFichier
' Rôle      : Affiche la boite de dialogue qui permet de sélectionner un fichier qui correspond aux filtres
' Paramètre : N/A
' Résultat  : renseigne le nom du fichier
' Erreurs   : N/A
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
' Méthode   : SelectionnerRepertoire
' Rôle      : Affiche la boite de dialogue qui permet de sélectionner un répertoire
' Paramètre : N/A
' Résultat  : renseigne le répertoire
' Erreurs   : N/A
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
' Méthode   : EstFichierVolumineux
' Rôle      : Indique si un fichier dépasse 4 Go (peut poser pb avec AdoDB.Stream)
' Propriété obligatoire :
'   - NomFichier
' Résultat  : Vrai si le fichier a une taille > 4 Go sinon Faux
' Erreurs   : -310 nom de fichier non renseigné
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
' Méthode   :
' Rôle      : Retourne la longueur d'un fichier (en octets)
' Paramètre : NomFichiser (string)
' Propriétés facultatives :
'   - NomFichier
' Résultat  : taille du fichier
' Erreurs   : -320 nom de fichier non renseigné
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
' Méthode   : FichierExiste
' Rôle      : Vérifie si un fichier existe
' Paramètre : NomFichier (string)
' Propriétés facultatives :
'   - NomFichier
' Résultat  : Vrai si le fichier physique existe sinon Faux
' Erreurs   : -330 nom du fichier non renseigné
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
' Méthode   : RepertoireExiste
' Rôle      : Vérifie si un répertoire existe
' Paramètre : Repertoire (string)
' Propriétés facultatives :
'   - Repertoire
' Résultat  : Vrai si le répertoire existe sinon Faux
' Erreurs   : -340 nom du répertoire non renseigné
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
' Méthode   : ConvertirUrlSharePoint
' Rôle      : Convertit un nom de chemin défini par une URL OneDrive ou SharePoint vers un nom de chemin Windows
' Paramètre : Chemin (string)
' Étapes    :
'   1. Convertit les URL de type Sharepoint
'   2. Convertit les URL des espaces de travail partagés
'   3. Convertit les URL de OneDrive personnel
' Résultat  : Chemin du fichier au format Windows
' Erreurs   : N/A
'-------------------------------------------------------------------------------------------------------------------------
Private Function ConvertirUrlSharePoint(Chemin As String) As String

    Dim sListeDossiers() As String, iNbDossiers As Integer, lPosDoc As Long, Repertoire As String
    
    ' Si le chemin du fichier commence par http
    If LCase$(Left(Chemin, 4)) = "http" Then
        Select Case True
        ' Espace personnel sur SharePoint (i.e. OneDrive Commercial)
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
' Méthode   : Extension
' Rôle      : Retourne l'extension du fichier
' Paramètre : N/A
' Résultat  : Extension du fichier
' Erreurs   : N/A
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
' Méthode   : Class_Initialize
' Rôle      : Création des objets ADODB.Stream et le dictionnaire des encodages.
' Paramètre : N/A
' Étapes    :
'   1. Création d'n objet AdoDB.stream
'   2. Ajout des enum des encodages dans un dictionnaire
'   3. Initialisation des compteurs
' Résultat  : Flux AdoDB.stream créé, liste des encodages prédéfinie
' Erreurs   : N/A
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
' Méthode   : Class_Terminate
' Rôle      : Libère proprement les objets internes.
'-------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()

    Set pFichierADODB = Nothing
    Set DicEncodage = Nothing

End Sub
