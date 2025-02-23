# VBA_Outils/ADODB
<h1>Licence</h1>
<p>Ce projet est distribué sous licence MIT. Consultez le fichier LICENSE pour plus de détails.</p>
<h1>Prérequis</h1>
<p>Environnement de développement : Microsoft Visual Basic for Applications (VBA)</p>
<h1>Automatisation des accès aux fichiers dans Excel</h1>
<p>Découvrez notre projet VBA conçu pour simplifier et automatiser les accès aux fichiers dans Microsoft Excel. Une solution robuste pour les développeurs VBA expérimentés cherchant à optimiser leur gestion des fichiers.</p>
<p>Notre classe encapsule des propriétés et des méthodes essentielles pour simplifier les opérations de traitement des fichiers. Elle offre une flexibilité pour définir les caractéristiques des fichiers et gérer les accès.</p>
<p>Les méthodes intégrées effectuent des contrôles en amont des actions, assurant l'intégrité et la sécurité des opérations sur les fichiers. Cette approche minimise les erreurs.</p>
<h2>Propriétés clés de la classe</h2>
<h3>Configuration</h3>
<ul>
<li><strong>TypeFichier</strong> : type du fichier (texte ou binaire)</li>
<li><strong>Encodage</strong> : encodage du fichier (enum)</li>
<li><strong>EncodageTxt</strong> : encodage du fichier (saisie libre)</li>
<li><strong>SeparateurLigne</strong> : séparateur de lignes d'un fichier texte</li>
</ul>
<h3>Gestion des accès</h3>
<ul>
<li><strong>TypeAcces</strong> : accès au fichier (lecture, écriture)</li>
<li><strong>NomFichier</strong> : nom du fichier (y compris le chemin d'accès)</li>
<li><strong>Fichier</strong>: permettre de pointer sur un objet ADODB Stream en dehors de la classe</li>
</ul>
<h3>Interface Utilisateur</h3>
<ul>
<li><strong>TitreboiteDeDialogue</strong> : titre affiché dans les boîtes de dialogue</li>
<li><strong>LibelleFiltre</strong> : libellé du filtre utilise pour afficher les fichiers dans une boîte de dialogue</li>
<li><strong>ExtensionFiltre</strong> : filtre des extensions de fichier utilise dans une boîte de dialogue</li>
<li><strong>NomInitialFichier</strong> : nom initial du fichier affiché dans une boîte de dialogue lors d'un enregistrement sous</li>
<li><strong>Filtre</strong> : filtre des extensions (parmi celles proposées par Excel) de fichier, utilisé dans une boîte de dialogue</li>
</ul>
<h3>Suivi des opérations</h3>
<ul>
<li><strong>NbreEnregLus</strong> : nombre d'enregistrements lus</li>
<li><strong>NbreEnregEcrits</strong> : nombre d'enregistrements écrits</li>
<li><strong>NbreOctectsLus</strong> : nombre d'octets lus</li>
<li><strong>NbreOctectsEcrits</strong> : nombre d'octets écrits</li>
</ul>
<h2>Méthodes essentielles pour la manipulation de fichiers</h2>
<h3>Méthodes Ouvrir et Fermer pour une gestion sécurisée des flux de données via ADODB.</h3>
<ul>
<li><strong>Ouvrir</strong> : ouvrir un fichier via ADODB</li>
<li><strong>Fermer</strong> : fermer un fichier</li>
</ul>
<h3>Lecture des données : extraire efficacement les données de vos fichiers.</h3>
<ul>
<li><strong>LireFichier</strong> : lire l'integralité d'un fichier (texte ou binaire)</li>
<li><strong>LireEnregistrement</strong> : lire le prochain enregistrement (jusqu'au prochain séparateur de lignes) d'un fichier texte</li>
<li><strong>FinFichier</strong> : fin du fichier texte atteinte</li>
<li><strong>Lire</strong> : lire n caractères d'un fichier (texte ou binaire)</li>
</ul>
<h3>Écriture de données : sauvegarder vos données de manière structurée.</h3>
<ul>
<li><strong>Ecrire</strong> : écrire des données dans un flux texte ou binaire</li>
<li><strong>EcrireEnregistrement</strong> : écrire un enregistrement dans un flux texte</li>
<li><strong>EnregistrerSous</strong> : enregistrer sous le nom du fichier le contenu du flux ADODB.Stream</li>
</ul>
<h3>Méthodes pour afficher des boîtes de dialogue intuitives pour la sélection de fichiers et répertoires.</h3>
<ul>
<li><strong>Repertoire</strong> : répertoire sélectionné dans une boîte de dialogue</li>
<li><strong>SelectionnerFichierEnregistrerSous</strong>: afficher la boîte de dialogue de sélection d'un fichier à enregsitrer sous (avec saisie du nom du fichier)</li>
<li><strong>SelectionnerFichier</strong> : afficher la boîte de dialogue de sélection d'un fichier</li>
<li><strong>SelectionnerRepertoire</strong> : afficher la boîte de dialogue de sélection d'un répertoire</li>
</ul>
<h2>Fonctionnalités de gestion des fichiers</h2>
<ul>
<li><strong>EstFichierVolumineux</strong> : vérifier si un fichier dépasse 4 Go (taille maximale des fichiers texte via ADODB.stream)</li>
<li><strong>FichierExiste</strong> : vérifier si le fichier dont le nom est "NomFichier" existe</li>
<li><strong>RepertoireExiste</strong> : vérifier si le répertoire dont le nom est "NomRepertoire" existe</li>
<li><strong>LongueurFichier</strong> : retourner la longueur d'un fichier en octets</li>
</ul>
