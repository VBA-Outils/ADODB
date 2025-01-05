# Outils_VBA/ADODB
<h1>Automatisation des Accès aux Fichiers dans Excel</h1>
<p>Découvrez notre projet VBA conçu pour simplifier et automatiser les accès aux fichiers dans Microsoft Excel. Une solution robuste pour les développeurs VBA expérimentés cherchant à optimiser leur gestion des fichiers.</p>
<p>Notre classe principale encapsule des propriétés et des méthodes essentielles pour simplifier les opérations de traitement des fichiers. Elle offre une flexibilité inégalée pour définir les caractéristiques physiques des fichiers et gérer les accès.</p>
<p>Les méthodes intégrées effectuent des contrôles en amont des actions, assurant l'intégrité et la sécurité des opérations sur les fichiers. Cette approche minimise les erreurs.</p>
<h2>Propriétés Clés de la Classe</h2>
<h3>Configuration Flexible</h3>
<ul>
<li>TypeFichier : type du fichier (texte ou binaire)</li>
<li>Encodage : encodage du fichier (enum)</li>
<li>EncodageTxt : encodage du fichier (saisie libre)</li>
<li>SeparateurLigne : séparateur de lignes d'un fichier texte</li>
</ul>
<h3>Gestion des Accès</h3>
<ul>
<li><font color="#FF4500">TypeAcces</font> : accès au fichier (lecture, écriture)</li>
<li>NomFichier : nom du fichier (y compris le chemin d'accès)</li>
<li>Fichier: permettre de pointer sur un objet ADODB Stream en dehors de la classe</li>
</ul>
<h3>Interface Utilisateur</h3>
<ul>
<li>TitreboîteDeDialogue : titre affiché dans les boîtes de dialogue</li>
<li>LibelléFiltre : libellé du filtre utilise pour afficher les fichiers dans une boîte de dialogue</li>
<li>ExtensionFiltre : filtre des extensions de fichier utilise dans une boîte de dialogue</li>
<li>NomInitialFichier : nom initial du fichier affiché dans une boîte de dialogue lors d'un enregistrement  sous</li>
<li>Filtre : filtre des extensions (parmi celles proposées par Excel) de fichier utilise dans une boîte de dialogue</li>
</ul>
<h3>Suivi des Opérations</h3>
<ul>
<li>NbreEnregLus : nombre d'enregistrements lus</li>
<li>NbreEnregEcrits : nombre d'enregistrements écrits</li>
<li>NbreOctectsLus : nombre d'octets lus</li>
<li>NbreOctectsEcrits : nombre d'octets écrits</li>
</ul>
<h2>Méthodes essentielles pour la manipulation de fichiers</h2>
<h3>Méthodes Ouvrir et Fermer pour une gestion sécurisée des flux de données via ADODB.</h3>
<ul>
<li>Ouvrir : ouvrir un fichier via ADODB</li>
<li>Fermer : fermer un fichier</li>
</ul>
<h3>Lecture des données : extraire efficacement les données de vos fichiers.</h3>
<ul>
<li>LireFichier : lire l'integralité d'un fichier (texte ou binaire)</li>
<li>LireEnregistrement : lire le prochain enregistrement (jusqu'au prochain séparateur de lignes) d'un fichier texte</li>
<li>FinFichier : fin du fichier texte atteinte</li>
<li>Lire : lire n caractères d'un fichier (texte ou binaire)</li>
</ul>
<h3>Écriture de données : sauvegarder vos données de manière structurée.</h3>
<ul>
<li>Ecrire : écrire des données dans un fichier texte ou binaire</li>
<li>EcrireEnregistrement : écrire un enregistrement dans un fichier texte</li>
<li>EnregistrerSous : enregistrer sous le nom du fichier le contenu de l'objet ADODB.Stream</li>
</ul>
<h3>Méthodes pour afficher des boîtes de dialogue intuitives pour la sélection de fichiers et répertoires.</h3>
<ul>
<li>Repertoire : répertoire sélectionné dans une boîte de dialogue</li>
<li>SelectionnerFichierEnregistrerSous: afficher la boîte de dialogue de sélection d'un fichier à enregsitrer sous (avec saisie du nom du fichier)</li>
<li>SelectionnerFichier : afficher la boîte de dialogue de sélection d'un fichier</li>
<li>SelectionnerRepertoire : afficher la boîte de dialogue de sélection d'un répertoire</li>
</ul>
<h2>Fonctionnalités de gestion des fichiers</h2>
<ul>
<li>EstFichierVolumineux : vérifier si un fichier dépasse 4 Go (taille maximale des fichiers texte via ADODB.stream)</li>
<li>FichierExiste : vérifier si le fichier dont le nom est "NomFichier" existe</li>
<li>RepertoireExiste : vérifier si le répertoire dont le nom est "NomRepertoire" existe</li>
<li>LongueurFichier : retourner la longueur d'un fichier en octets</li>
</ul>
