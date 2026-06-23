# 📁 ADODB – Classe VBA avancée pour la gestion des fichiers texte et binaires

![VBA](https://img.shields.io/badge/VBA-Excel-blue)
![License](https://img.shields.io/badge/License-MIT-green)

Une classe VBA complète et robuste permettant de lire, écrire et manipuler des fichiers texte ou binaires via ADODB.Stream.
Elle offre une interface simple, cohérente et sécurisée pour gérer les fichiers dans vos projets VBA, tout en intégrant des boîtes de dialogue natives et un suivi précis des opérations.

Limites connues : lors de la lecture de fichiers texte, dont la taille est proche de ou dépasse 4 Go, des pertes de données ont été constatées. Certaines parties du fichier sont ignorées par AdoDB.Stream, et par conséquent, ne sont pas restituées par la méthode LireEnregistrement. Ces enregistrements ignorés restent négligeables par rapport à la taille du fichier, mais induisent un manque de fiabilité dans la gestion des fichiers volumineux.

---

## ⭐ Pourquoi utiliser cette classe ?

### ✔️ Une alternative moderne aux fonctions VBA natives
Les méthodes `Open`, `Input`, `Line Input`, `Print` sont limitées et peu fiables selon l’encodage.
`ADODB.Stream` offre une gestion plus stable, plus rapide et compatible avec les encodages modernes.

### ✔️ Gestion automatique des encodages
UTF‑8, UTF‑16, ANSI, binaire…
La classe encapsule toute la complexité et vous permet de choisir simplement l’encodage souhaité.

### ✔️ Interface utilisateur intégrée
Sélection de fichiers, enregistrement sous, choix de répertoire…
Sans API Windows ni code complexe.

### ✔️ API unifiée
Même logique pour le texte et le binaire.
Chaque propriété est définie avec un Enum afin de disposer de tous les choix possibles en clair (sans devoir consulter la documentation technique).
Plus besoin de jongler entre plusieurs syntaxes VBA.

### ✔️ Sécurité et robustesse
- Vérification d’existence des fichiers et répertoires
- Gestion propre des flux
- Compteurs d’octets et d’enregistrements
- Détection de fin de fichier

### ✔️ Idéal pour les projets professionnels
La classe est pensée pour être :
- réutilisable
- stable
- documentée
- facile à intégrer

---

## ✨ Fonctionnalités principales

### 🔧 Configuration du fichier
- TypeFichier : texte ou binaire
- Encodage / EncodageTxt : choix de l'encodage avec une Enum ou saisie libre
- SeparateurLigne : séparateur de lignes (CR, LC, CR/LF)

### 🔒 Gestion des accès
- ModeAcces : lecture, modification ou écriture
- NomFichier : nom du fichier physique à traiter
- Ouvrir : Ouvrir un fichier via ADODB
- LireFichier : Lire l'intégralité d'un fichier (texte ou binaire)
- LireEnregistrement : Lire un fichier texte jusqu'au prochain séparateur
- FinFichier : Fin du fichier texte atteinte
- Lire : Lire n caractères d'un fichier (texte ou binaire)
- Ecrire : Ecrire des données dans un fichier texte ou binaire
- EcrireEnregistrement : Ecrire un enregistrement dans un fichier texte
- EnregistrerSous : Enregistrer sous le nom de fichier le contenu de l'objet ADODB.Stream
- Fermer : Fermer un fichier
 

### 🖥️ Interface utilisateur
- TitreBoiteDeDialogue : Titre affiché dans les boîtes de dialogue
- LibelleFiltre : Libellé du filtre utilisé pour afficher les fichiers dans une boîte de dialogue
- ExtensionFiltre : Filtre des extensions de fichier utilisé dans une boîte de dialogue
- NomInitialFichier : Nom initial du fichier affiché dans une boîte de dialogue lors d'un enregsitrement sous
- Filtre : Filtre des extensions (parmi celles proposées par Excel) de fichier utilisé dans une boîte de dialogue
- SelectionnerFichier : sélectionner le nom d'un fichier à lire dans une boite de dialogue Excel
- SelectionnerFichierEnregistrerSous : saisir le nom d'un fichier à écrire dans une boite de dialogue Excel
- SelectionnerRepertoire : sélectionner un répertoire à traiter

### 📊 Suivi des opérations
- NbreEnregLus / NbreEnregEcrits
- NbreOctetsLus / NbreOctetsEcrits

### 🧪 Utilitaires
- FichierExiste
- RepertoireExiste
- EstFichierVolumineux
- LongueurFichier

---

## 🧩 Dépendances

### Classes / Modules nécessaires

- `AdoDB.cls` (cette classe)
- `AdoDB_Enum.bas` (module contenant les Enum)
- Référence obligatoire :  
  **Microsoft ActiveX Data Objects x.x Library**

### Objets utilisés

- `ADODB.Stream`
- `Scripting.Dictionary`

---

## ⚙️ Installation

1. Importer les fichiers :
   - `AdoDB.cls`
   - `AdoDB_Enum.bas`

2. Activer la référence :
   - **Microsoft ActiveX Data Objects 6.1 Library** (ou version équivalente)

3. Vérifier que les Enum sont bien dans un module standard.

---

## 🚀 Exemples d’utilisation

### Ecriture d'un fichier texte

```vba
    ' Déclaration d'un objet ADODB
    Dim oFichier As New ADODB
    ' Enregistrement
    Dim sEnreg As String
    
    ' Créer un fichier texte
    ' Ecriture dans le flux ADODB de l'entête du fichier texte
    sEnreg = "Marque;Modele;Categorie;Carburant;Puissance" & vbCrLf
    
    ' Ouverture du fichier CSV en écriture
    With oFichier
        .TypeFichier = AD_TYPE_TEXT
        .TypeAcces = AD_MODE_WRITE
        .Encodage = AD_UTF_8
        .Ouvrir
        ' Ecriture des donnees dans le flux ADODB
        .EcrireEnregistrement (sEnreg)
        ' Sélectionner le nom du fichier
        .NomInitialFichier = Environ("OneDrive") & "\Documents\tests.txt"
        .ExtensionFiltre = "*.txt"
        .TitreBoiteDeDialogue = "Enregistrer le fichier sous"
        .SelectionnerFichierEnregistrerSous
        ' Si le nom du fichier n'a pas été sélectionné alors on quitte
        If .NomFichier <> "" Then 
            ' Enregistrement du flux ADODB
            .EnregistrerSous
            ' Fermeture du flux
        End If
        .Fermer
    End With

    ' Libérer les ressources
    Set oFichier = Nothing

```

### Lecture d'un fichier texte

```vba
    ' Déclaration d'un objet ADODB
    Dim oFichier As New ADODB
    ' Enregistrement
    Dim sEnreg As String

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
```

---

## 📦 Structure du projet

```
ADODB/
 ├── ADODB.cls
 ├── AdoDB_Enum.bas
 ├── ExempleADODB.bas
 ├── LICENSE
 └── README.md
```

---

## 🛠️ Prérequis

- Microsoft Excel / VBA  
- Référence Microsoft ActiveX Data Objects x.x Library

---

## 📄 Licence

Projet distribué sous licence MIT.

---

## 🤝 Contribution

Les contributions sont les bienvenues :  
- suggestions  
- corrections  
- nouvelles fonctionnalités  

Ouvrez une issue ou une pull request.
