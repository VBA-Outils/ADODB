# 📁 ADODB – Classe VBA avancée pour la gestion des fichiers texte et binaires

![VBA](https://img.shields.io/badge/VBA-Excel-blue)
![License](https://img.shields.io/badge/License-MIT-green)

Une classe VBA complète et robuste permettant de lire, écrire et manipuler des fichiers texte ou binaires via ADODB.Stream.  
Elle offre une interface simple, cohérente et sécurisée pour gérer les fichiers dans vos projets VBA, tout en intégrant des boîtes de dialogue natives et un suivi précis des opérations.

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

### 🖥️ Interface utilisateur
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

## 🚀 Exemples d’utilisation

### Ecriture d'un fichier texte

```vba
    ' Déclaration d'un objet ADODB
    Dim oFichier As New ADODB
    ' Enregistrement et nom du fichier
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
