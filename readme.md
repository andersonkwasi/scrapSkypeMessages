# Skype Web Extractor

## Description
Ce projet contient un script Python qui utilise Selenium pour extraire automatiquement les messages de conversations Skype via l'interface web de Skype. Il se connecte à votre compte Skype, parcourt toutes les conversations visibles, extrait les messages et les exporte dans un fichier Excel.

## Fonctionnalités
- Connexion à Skype Web via Selenium (l'utilisateur saisit ses identifiants manuellement sur l'interface Skype).
- Extraction d'un nombre limité de conversations (par défaut : 10).
- Extraction des messages sans duplication grâce à un identifiant unique basé sur l'expéditeur, le contenu et l'heure.
- Exportation des messages extraits au format Excel.

## Prérequis
- Python 3.6+
- Chrome ou Chromium
- ChromeDriver (assurez-vous qu'il correspond à la version de votre Chrome)

## Installation

1. Clonez ce dépôt ou téléchargez le script `skype_web_extractor.py`.

2. Installez les dépendances nécessaires :
   ```
   pip install selenium pandas tqdm openpyxl
   ```

3. Téléchargez ChromeDriver depuis [le site officiel](https://sites.google.com/a/chromium.org/chromedriver/downloads) et placez-le dans votre PATH ou dans le même dossier que le script.


## Utilisation

1. Ouvrez un terminal et naviguez jusqu'au dossier contenant le script.

2. Exécutez le script :
   ```
   python3 skype_web_extractor.py
   ```

3. Le script va :
   - Ouvrir une fenêtre Chrome automatiquement une fenêtre Chrome et se dirigera vers la page de connexion Skype Web.
   - Saisissez vos identifiants Skype dans l'interface, puis attendez que les messages se chargent.
   - Parcourir toutes les conversations visibles
   - Extraire les messages
   - Exporter les messages dans un fichier CSV nommé `skype_messages_YYYYMMDD_HHMMSS.csv`

4. Une fois l'exécution terminée, vous trouverez le fichier Excel dans le même dossier que le script.

5. Extraction des conversations :

    Par défaut, le script extrait jusqu'à 10 conversations. Ce paramètre peut être modifié dans la fonction extract_conversations en changeant la valeur de limit.


![alt text](<Capture d’écran du 2024-10-18 00-28-17.png>)


## Remarques

- Le script peut prendre un certain temps à s'exécuter, surtout si vous avez de nombreuses conversations ou messages.
- Assurez-vous d'avoir une connexion Internet stable pendant l'exécution du script.
- Le script peut nécessiter des ajustements en fonction des mises à jour de l'interface Skype Web.

## Dépannage

- Si le script échoue à se connecter, vérifiez vos identifiants et assurez-vous que vous pouvez vous connecter manuellement à Skype Web.
- Si le script ne parvient pas à trouver certains éléments, il peut être nécessaire d'ajuster les sélecteurs CSS ou XPath utilisés dans le code.

## Contribution

Les contributions à ce projet sont les bienvenues. N'hésitez pas à ouvrir une issue ou à soumettre une pull request.

## Licence

Ce projet est sous licence MIT. Voir le fichier `LICENSE` pour plus de détails.
