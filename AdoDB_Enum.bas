Attribute VB_Name = "AdoDB_Enum"
'
' https://github.com/VBA-Outils/
'
' Fonctions génériques VBA
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

' Propriété "SeparateurLigne" : caractère(s) de fin d'enregistrement
Public Enum LineSeparatorsEnum
    AD_CR_LF = -1          ' Par défaut, Retour Chariot et Ligne Suivante
    AD_LF = 10             ' Ligne Suivante
    AD_CR = 13             ' Retour Chariot
End Enum

' Propriété "encodage" : encodage du fichier
' La liste des encodages n'est pas exhaustive et doit être complétée si nécessaire. Les plus utilisés sont présents.
Public Enum CharsetEnum
    AD_ASCII = 1
    AD_LATIN_2 = 2
    AD_LATIN_4 = 3
    AD_CYRILLIQUE = 4
    AD_GREC = 5
    AD_LATIN_5 = 6
    AD_COREEN = 7
    AD_UTF_7 = 8
    AD_UTF_8 = 9
    AD_UTF_8_BOM = 10
    AD_CP1250 = 11
    AD_CP1251 = 12
    AD_X_ANSI = 13
    AD_ISO_8859_11 = 14
    AD_ARABE = 15
    AD_AUTRE = 16
End Enum

' Proprieté "TypeFichier" : type de format du fichier
Public Enum StreamTypeEnum
    AD_TYPE_BINARY = 1         ' Fichier Binaire
    AD_TYPE_TEXT = 2           ' Fichier Texte
End Enum

' Propriété "TypeAcces"
Public Enum ConnectModeEnum
    AD_MODE_READ = 1         ' Ouverture en lecture du fichier
    AD_MODE_WRITE = 2        ' Ouverture en écriture du fichier
    AD_MODE_READ_WRITE = 3   ' Ouverture en lecture / écriture du fichier
End Enum

' Filtres pour la boîte de dialogue "Enregistrer sous" (les plus utilisés)
Public Enum FilterIndexEnum
    AD_XLSX = 1            ' Classeur Excel
    AD_XLSM = 2            ' Classeur Excel prenant en charge les macros
    AD_XLSB = 3            ' Classeur Excel binaire
    AD_XLS = 4             ' Classeur Excel 97-2003
    AD_CSV_COMA = 5        ' CSV UTF-8 (délimité par des virgules)
    AD_XML = 6             ' Données XML
    AD_XLTX = 9            ' Modèle Excel
    AD_XLTM = 10           ' Modèle Excel prenant en charge les macros
    AD_CSV_TAB = 12        ' CSV (séparateur : tabulation)
    AD_TXT_UNICODE = 13    ' Texte unicode
    AD_CSV_SEMICOLON = 16  ' CSV (séparateur : point-virgule)
End Enum

' Propriété de la méthode "Lire"
Public Enum StreamReadEnum
    AD_READ_ALL = -1       ' Lit l'intégralité du fichier. C’est la seule valeur valide s'il s'agit d'un fichier binaire.
    AD_READ_LINE = -2      ' Lit l'enregistrement suivant (jusqu'au prochain séparateur de lignes ou la fin de fichier).
End Enum

Public Enum SaveOptionsEnum
    AD_SAVE_CREATE_NOT_EXIST = 1
    AD_SAVE_CREATE_OVER_WRITE = 2
End Enum
