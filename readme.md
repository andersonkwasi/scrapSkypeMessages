# Skype Web Message Extractor

## Description

Le **Skype Web Message Extractor** est un script Python automatisé qui utilise Selenium pour extraire des messages à partir de conversations Skype Web. Le script permet de se connecter à Skype Web, d'extraire un nombre limité de conversations et d'exporter les messages au format Excel (.xlsx) pour analyse et archivage.

### Fonctionnalités :
- Connexion à Skype Web via Selenium (l'utilisateur saisit ses identifiants manuellement sur l'interface Skype).
- Extraction d'un nombre limité de conversations (par défaut : 10).
- Extraction des messages sans duplication grâce à un identifiant unique basé sur l'expéditeur, le contenu et l'heure.
- Exportation des messages extraits au format Excel.

## Prérequis

Avant de pouvoir utiliser ce script, assurez-vous que les logiciels et modules suivants sont installés sur votre machine :

- **Python 3.x** (disponible ici : https://www.python.org/downloads/)
- **Google Chrome** (le script utilise le navigateur Chrome)
- **ChromeDriver** : Un composant essentiel qui permet à Selenium de contrôler Chrome. Assurez-vous que la version de ChromeDriver correspond à celle de Chrome installée sur votre machine. Vous pouvez télécharger ChromeDriver ici : https://sites.google.com/a/chromium.org/chromedriver/downloads

Vous pouvez vérifier la version de Chrome installée en utilisant cette commande :

```bash
google-chrome --version


- ** Installation des modules Python nécessaires :**

Le script utilise plusieurs modules Python, notamment Selenium, pandas et tqdm. Vous pouvez installer ces modules en utilisant pip :

bash

 pip install selenium pandas tqdm openpyxl

openpyxl est nécessaire pour l'exportation en format Excel.
Installation

    Téléchargement de ChromeDriver :
        Téléchargez la version correspondante de ChromeDriver à partir de ChromeDriver.
        Extrayez le fichier téléchargé et placez-le dans un répertoire de votre choix.
        Ajoutez le répertoire contenant chromedriver à votre PATH système, ou placez-le dans le répertoire où vous allez exécuter le script.

    Configuration du projet :
        Clonez ce projet ou téléchargez les fichiers sources.
        Assurez-vous d'avoir Python 3.x et les modules requis installés.

Utilisation

    Exécution du script :
        Ouvrez un terminal dans le répertoire contenant le script skype.py.
        Exécutez le script avec la commande suivante :

    bash

    python skype.py

    Connexion :
        Le script ouvrira automatiquement une fenêtre Chrome et se dirigera vers la page de connexion Skype Web.
        Saisissez vos identifiants Skype dans l'interface, puis attendez que les messages se chargent.
        Une fois connecté et les messages chargés, appuyez sur Entrée dans le terminal pour commencer l'extraction des conversations.

    Extraction des conversations :
        Par défaut, le script extrait jusqu'à 10 conversations. Ce paramètre peut être modifié dans la fonction extract_conversations en changeant la valeur de limit.
        Les messages sont extraits sans duplication, même si le même message apparaît plusieurs fois dans différentes parties de la conversation.

    Exportation des messages :
        Les messages extraits sont automatiquement exportés dans un fichier Excel (.xlsx) nommé selon la date et l'heure actuelles, par exemple : skype_messages_20241017_232854.xlsx.
        Le fichier est généré dans le même répertoire où le script est exécuté.

Exemple d'exportation

Voici un exemple de format de fichier exporté :
Nom	Message	Heure
John Doe	Hello! How are you?	12:45 PM
Jane Doe	I'm good, thank you!	12:46 PM
John Doe	Let's meet tomorrow at 10 AM.	12:50 PM
Journalisation

Le script enregistre les actions et les erreurs dans un fichier de log nommé skype_extractor.log. Ce fichier peut être utilisé pour déboguer ou vérifier l'état du processus d'extraction.
Remarques

    Le script est conçu pour fonctionner avec Google Chrome et Skype Web. Si vous utilisez un autre navigateur, des modifications seront nécessaires.
    Si une conversation contient de nombreux messages, cela peut prendre un certain temps pour les charger tous. Soyez patient.
    Si vous avez des erreurs avec ChromeDriver, assurez-vous que la version de ChromeDriver correspond à la version de Google Chrome installée sur votre machine.

Limitations

    Le script extrait actuellement un nombre limité de conversations (par défaut 10). Cette limite peut être modifiée dans le code en ajustant le paramètre limit.
    Le script ne gère pas la pagination automatique pour les conversations qui s'étendent sur plusieurs pages dans l'interface Skype Web.

Améliorations futures

    Ajouter la gestion de la pagination pour extraire des messages de très longues conversations.
    Ajouter des options de ligne de commande pour définir le nombre de conversations à extraire ou le format d'exportation.
    Supporter d'autres formats d'exportation comme JSON ou CSV.

Auteur

Ce script a été développé par AndersonKwsy pour faciliter l'extraction de messages depuis Skype Web.

N'hésitez pas à me contacter pour toute question ou assistance concernant ce projet.