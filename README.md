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
- Encodage / EncodageTxt  
- SeparateurLigne  

### 🔒 Gestion des accès
- TypeAcces : lecture ou écriture  
- NomFichier  
- Fichier : accès direct à ADODB.Stream  

### 🖥️ Interface utilisateur
- SelectionnerFichier  
- SelectionnerFichierEnregistrerSous  
- SelectionnerRepertoire  

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

### Ecriture puis lecture d'un fichier texte

```vba
    ' Déclaration d'un objet ADODB
    Dim oFichier As New ADODB
    ' Enregistrement et nom du fichier
    Dim sEnreg As String, sNomFichier As String
    
    ' Créer un fichier texte
    
    ' Ouverture du fichier CSV en écriture
    With oFichier
        .TypeFichier = AD_TYPE_TEXT
        .TypeAcces = AD_MODE_WRITE
        .Encodage = AD_UTF_8
        .Ouvrir
        ' Ecriture dans le flux ADODB de l'entête du fichier texte
        sEnreg = "Marque;Modele;Categorie;Carburant;Puissance" & vbCrLf
        .EcrireEnregistrement (sEnreg)
        ' Ecriture des donnees dans le flux ADODB
        sEnreg = "Marque1;Modèle1;SUV;Essence;129 ch" & vbCrLf
        .EcrireEnregistrement (sEnreg)
        sEnreg = "Marque2;Modèle2;SUV;Essence;89 ch" & vbCrLf
        .EcrireEnregistrement (sEnreg)
        ' Sélectionner le nom du fichier
        .NomInitialFichier = Environ("OneDrive") & "\Documents\tests.txt"
        .ExtensionFiltre = "*.txt"
        .TitreBoiteDeDialogue = "Enregistrer le fichier sous"
        .SelectionnerFichierEnregistrerSous
        ' Si le nom du fichier n'a pas été sélectionné alors on quitte
        If .NomFichier = "" Then Exit Sub
        ' Récupère le nom du fichier afin de pouvoir le lire dans l'étape suivante
        sNomFichier = .NomFichier
        ' Enregistrement du flux ADODB
        .EnregistrerSous
        ' Fermeture du flux
        .Fermer
    End With
    
    ' Lire le fichier texte précédemment créé
    
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
