# Outils_VBA/ADODB

 Classe VBA qui permet de lire et écrire des fichiers (texte ou binaire) via les flux ADODB.<br>
 Les propriétés de la classe permettent de définir les caractéristiques des fichiers, les méthodes de gérer les fichiers (ouverture, lecture/écriture, fermeture).<br>
 Des boîtes de dialogue sont proposées afin de pouvoir sélectionner un répertoire, un fichier.<br>
<br>
 Proprietes de la classe :<br>
-------------------------<br>
  TypeFichier                        : Type du fichier (texte ou binaire)<br>
  Encodage                           : Encodage du fichier (enum)<br>
  EncodageTxt                        : Encodage du fichier (saisie libre)<br>
  SeparateurLigne                    : Separateur de lignes d'un fichier texte<br>
  TypeAcces                          : Acces au fichier (lecture, ecriture)<br>
  TitreBoiteDeDialogue               : Titre affiche dans les boites de dialogue<br>
  LibelleFiltre                      : Libelle du filtre utilise pour afficher les fichiers dans une boite de dialogue<br>
  ExtensionFiltre                    : Filtre des extensions de fichier utilise dans une boite de dialogue<br>
  NomInitialFichier                  : Nom initial du fichier affiche dans une boite de dialogue lors d'un enregsitrement sous<br>
  Filtre                             : Filtre des extensions (parmi celles proposees par Excel) de fichier utilise dans une boite de dialogue<br>
  NomFichier                         : nom du fichier (y compris le chemin d'acces)<br>
  Repertoire                         : Repertoire selectionne dans une boite de dialogue<br>
  Fichier                            : Permettre de pointer sur un objet ADODB Stream en dehors de la classe<br>
  NbreEnregLus                       : Nombre d'enregistrements lus<br>
  NbreEnregEcrits                    : Nombre d'enregistrements ecrits<br>
  NbreOctectsLus                     : Nombre d'octets lus<br>
  NbreOctectsEcrits                  : Nombre d'octets ecrits<br>
<br>
Methodes de la classe :<br>
-----------------------<br>
  Ouvrir                             : Ouvrir un fichier via ADODB<br>
  LireFichier                        : Lire l'integralite d'un fichier (texte ou binaire)<br>
  LireEnregistrement                 : Lire un fichier texte jusqu'au prochain separateur<br>
  FinFichier                         : Fin du fichier texte atteinte<br>
  Lire                               : Lire n caracteres d'un fichier (texte ou binaire)<br>
  LireEnregistrement                 : Lire le prochain enregistrement (jusqu'au prochain separateur de lignes) d'un fichier texte<br>
  LireFichier                        : Lire l'integralite d'un fichier texte ou binaire<br>
  Fermer                             : Fermer un fichier<br>
  Ecrire                             : Ecrire des donnees dans un fichier texte ou binaire<br>
  EcrireEnregistrement               : Ecrire un enregistrement dans un fichier texte<br>
  EnregistrerSous                    : Enregistrer sous le nom de fichier le contenu de l'objet ADODB.Stream<br>
  SelectionnerFichierEnregistrerSous : Afficher la boite de dialogue de selection d'un fichier a enregsitrer sous (avec saisie du nom du fichier)<br>
  SelectionnerFichier                : Afficher la boite de dialogue de selection d'un fichier<br>
  SelectionnerRepertoire             : Afficher la boite de dialogue de selection d'un repertoire<br>
  EstFichierVolumineux               : Verifier si un fichier depasse 4 Go (taille maximale des fichiers texte via ADODB.stream)<br>
  FichierExiste                      : Verifier si le fichier dont le nom est "NomFichier" existe<br>
  RepertoireExiste                   : Verifier si le repertoire dont le nom est "NomRepertoire" existe<br>
  LongueurFichier                    : Retourner la longueur d'un fichier en octets<br>
<br>
Exemples d'utilisation de la classe ADODB présents dans les fichiers ExempleADODB joints.
