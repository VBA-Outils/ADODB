# Outils_VBA/ADODB

<h1>Outils-VBA/ADODB : Automatisation des Accès aux Fichiers dans Excel</h1>
<p>Découvrez notre projet VBA conçu pour simplifier et automatiser les accès aux fichiers dans Microsoft Excel. Une solution robuste pour les développeurs VBA expérimentés cherchant à optimiser leur gestion des fichiers.</p>
<p>Notre classe principale encapsule des propriétés et des méthodes essentielles pour simplifier les opérations de traitement des fichiers. Elle offre une flexibilité inégalée pour définir les caractéristiques physiques des fichiers et gérer les accès.</p>
<p>Les méthodes intégrées effectuent des contrôles en amont des actions, assurant l'intégrité et la sécurité des opérations sur les fichiers. Cette approche minimise les erreurs.</p>
<h2>Propriétés Clés de la Classe</h2>
<h3>Configuration Flexible</h3>
<ul>
<li>TypeFichier : Type du fichier (texte ou binaire)</li>
<li>Encodage : Encodage du fichier (enum)</li>
<li>EncodageTxt : Encodage du fichier (saisie libre)</li>
<li>SeparateurLigne : Separateur de lignes d'un fichier texte</li>
</ul>
<h3>Gestion des Accès</h3>
<ul>
<li>TypeAcces : Acces au fichier (lecture, ecriture)</li>
<li>NomFichier : nom du fichier (y compris le chemin d'acces)</li>
<li>Repertoire : Repertoire selectionne dune boite de dialogue</li>
<li>Fichier: Permettre de pointer sur un objet ADODB Stream en dehors de la classe</li>
</ul>
<h3>Interface Utilisateur</h3>
<ul>
<li>TitreBoiteDeDialogue : Titre affiche dans les boites de dialogue</li>
<li>LibelleFiltre : Libelle du filtre utilise pour afficher les fichiers dans une boite de dialogue</li>
<li>ExtensionFiltre : Filtre des extensions de fichier utilise dans une boite de dialogue</li>
<li>NomInitialFichier : Nom initial du fichier affiche dans une boite de dialogue lors d'un enregsitrement sous</li>
<li>Filtre : Filtre des extensions (parmi celles proposees par Excel) de fichier utilise dans une boite de dialogue</li>
</ul>
<h3>Suivi des Opérations</h3>
<ul>
<li>NbreEnregLus : Nombre d'enregistrements lus</li>
<li>NbreEnregEcrits : Nombre d'enregistrements ecrits</li>
<li>NbreOctectsLus : Nombre d'octets lus</li>
<li>NbreOctectsEcrits : Nombre d'octets ecrits</li>
</ul>
<h2>Méthodes Essentielles pour la Manipulation de Fichiers</h2>
<h3>Méthodes Ouvrir et Fermer pour une gestion sécurisée des flux de données via ADODB.</h3>
<ul>
<li>Ouvrir : Ouvrir un fichier via ADODB</li>
<li>Fermer : Fermer un fichier</li>
</ul>
<h3>Lecture des données : extraire efficacement les données de vos fichiers.</h3>
<ul>
<li>LireFichier : Lire l'integralite d'un fichier (texte ou binaire)</li>
<li>LireEnregistrement : Lire le prochain enregistrement (jusqu'au prochain separateur de lignes) d'un fichier texte</li>
<li>FinFichier : Fin du fichier texte atteinte</li>
<li>Lire : Lire n caracteres d'un fichier (texte ou binaire)</li>
</ul>
<h3>Écriture de données : sauvegarder vos données de manière structurée.</h3>
<ul>
<li>Ecrire : Ecrire des donnees dans un fichier texte ou binaire</li>
<li>EcrireEnregistrement : Ecrire un enregistrement dans un fichier texte</li>
<li>EnregistrerSous : Enregistrer sous le nom de fichier le contenu de l'objet ADODB.Stream</li>
</ul>
<h3>Méthodes pour afficher des boîtes de dialogue intuitives pour la sélection de fichiers et répertoires.</h3>
<ul>
<li>SelectionnerFichierEnregistrerSous: Afficher la boite de dialogue de selection d'un fichier a enregsitrer sous (avec saisie du nom du fichier)</li>
<li>SelectionnerFichier : Afficher la boite de dialogue de selection d'un fichier</li>
<li>SelectionnerRepertoire : Afficher la boite de dialogue de selection d'un repertoire</li>
</ul>
<h2>Fonctionnalités de gestion des fichiers</h2>
<ul>
<li>EstFichierVolumineux : Verifier si un fichier depasse 4 Go (taille maximale des fichiers texte via ADODB.stream)</li>
<li>FichierExiste : Verifier si le fichier dont le nom est "NomFichier" existe</li>
<li>RepertoireExiste : Verifier si le repertoire dont le nom est "NomRepertoire" existe</li>
<li>LongueurFichier : Retourner la longueur d'un fichier en octets</li>
</ul>
