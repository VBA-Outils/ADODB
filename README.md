# Outils_VBA/ADODB

 Classe VBA qui permait de lire et écrire des fichiers (texte ou binaire) via les flux ADODB.<br>
 Les propriétés de la classe permettent de définir les caractéristiques des fichiers, les méthodes de gérer les fichiers (ouverture, lecture/écriture, fermeture).<br>
 Des boîtes de dialogue sont proposées afin de pouvoir sélectionner un répertoire, un fichier.

 Proprietes de la classe :
-------------------------
  TypeFichier                        : Type du fichier (texte ou binaire)
  Encodage                           : Encodage du fichier (enum)
  EncodageTxt                        : Encodage du fichier (saisie libre)
  SeparateurLigne                    : Separateur de lignes d'un fichier texte
  TypeAcces                          : Acces au fichier (lecture, ecriture)
  TitreBoiteDeDialogue               : Titre affiche dans les boites de dialogue
  LibelleFiltre                      : Libelle du filtre utilise pour afficher les fichiers dans une boite de dialogue
  ExtensionFiltre                    : Filtre des extensions de fichier utilise dans une boite de dialogue
  NomInitialFichier                  : Nom initial du fichier affiche dans une boite de dialogue lors d'un enregsitrement sous
  Filtre                             : Filtre des extensions (parmi celles proposees par Excel) de fichier utilise dans une boite de dialogue
  NomFichier                         : nom du fichier (y compris le chemin d'acces)
  Repertoire                         : Repertoire selectionne dans une boite de dialogue
  Fichier                            : Permettre de pointer sur un objet ADODB Stream en dehors de la classe
  NbreEnregLus                       : Nombre d'enregistrements lus
  NbreEnregEcrits                    : Nombre d'enregistrements ecrits
  NbreOctectsLus                     : Nombre d'octets lus
  NbreOctectsEcrits                  : Nombre d'octets ecrits

Methodes de la classe :
-----------------------
  Ouvrir                             : Ouvrir un fichier via ADODB
  LireFichier                        : Lire l'integralite d'un fichier (texte ou binaire)
  LireEnregistrement                 : Lire un fichier texte jusqu'au prochain separateur
  FinFichier                         : Fin du fichier texte atteinte
  Lire                               : Lire n caracteres d'un fichier (texte ou binaire)
  LireEnregistrement                 : Lire le prochain enregistrement (jusqu'au prochain separateur de lignes) d'un fichier texte
  LireFichier                        : Lire l'integralite d'un fichier texte ou binaire
  Fermer                             : Fermer un fichier
  Ecrire                             : Ecrire des donnees dans un fichier texte ou binaire
  EcrireEnregistrement               : Ecrire un enregistrement dans un fichier texte
  EnregistrerSous                    : Enregistrer sous le nom de fichier le contenu de l'objet ADODB.Stream
  SelectionnerFichierEnregistrerSous : Afficher la boite de dialogue de selection d'un fichier a enregsitrer sous (avec saisie du nom du fichier)
  SelectionnerFichier                : Afficher la boite de dialogue de selection d'un fichier
  SelectionnerRepertoire             : Afficher la boite de dialogue de selection d'un repertoire
  EstFichierVolumineux               : Verifier si un fichier depasse 4 Go (taille maximale des fichiers texte via ADODB.stream)
  FichierExiste                      : Verifier si le fichier dont le nom est "NomFichier" existe
  RepertoireExiste                   : Verifier si le repertoire dont le nom est "NomRepertoire" existe
  LongueurFichier                    : Retourner la longueur d'un fichier en octets

Exemples d'utilisation de la classe ADODB présents dans les fichiers ExempleADODB joints.
