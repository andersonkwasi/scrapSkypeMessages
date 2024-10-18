from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
from datetime import datetime
import logging
from tqdm import tqdm

# Configurer la journalisation des erreurs
logging.basicConfig(filename='skype_extractor.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

class SkypeWebExtractor:
    """
    Classe pour extraire des messages depuis Skype Web à l'aide de Selenium.
    """

    def __init__(self):
        """
        Initialise le navigateur Chrome pour l'automatisation avec Selenium.
        """
        self.driver = webdriver.Chrome()
        self.messages = []
        self.unique_messages = set()  # Pour vérifier les doublons

    def login(self):
        """
        Effectue la connexion à Skype Web en attendant que l'utilisateur saisisse ses identifiants manuellement.
        """
        logging.info("Connexion à Skype Web.")
        try:
            self.driver.get("https://web.skype.com")
            WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located((By.NAME, "loginfmt"))
            )
            print("Veuillez saisir vos accès Skype.")
            input("Patientez que les messages se chargent puis tapez Entrée pour continuer...")
        except Exception as e:
            logging.error(f"Erreur lors de la tentative de connexion : {e}")
            raise e

    def extract_message_info(self, message_div):
        """
        Extrait les informations d'un message à partir d'un élément HTML.

        Args:
            message_div (WebElement): L'élément div contenant le message.

        Returns:
            tuple: Un tuple contenant le nom de l'expéditeur, le contenu du message et l'heure d'envoi.
        """
        aria_label = message_div.get_attribute("aria-label")
        parts = aria_label.split(", ")

        if len(parts) >= 3:
            name = parts[0]
            message = ", ".join(parts[1:-1])  # Le message peut contenir des virgules
            time = parts[-1].split(" à ")[-1]
            return name, message, time
        return None, None, None

    def extract_conversations(self, limit=10):
        """
        Extrait les conversations disponibles sur Skype Web, avec une limite sur le nombre de conversations à extraire.

        Args:
            limit (int): Le nombre maximal de conversations à extraire (par défaut 10).
        """
        logging.info("Recherche des conversations...")
        try:
            # Localisation des conversations
            conversation_list = WebDriverWait(self.driver, 30).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "[role='listitem']"))
            )

            logging.info(f"Nombre total de conversations trouvées : {len(conversation_list)}")
            print(f"Nombre total de conversations trouvées : {len(conversation_list)}")

            for i, conv in enumerate(tqdm(conversation_list[:limit], desc="Extraction des conversations"), 1):
                try:
                    logging.info(f"Extraction conversation {i}/{len(conversation_list)}")
                    conv.click()

                    # Attendre que les messages soient visibles
                    WebDriverWait(self.driver, 30).until(
                        EC.presence_of_element_located((By.XPATH, "//div[@role='region']"))
                    )

                    # Défilement pour charger les messages récents
                    previous_message_count = 0
                    while True:
                        # Attendre un peu pour le chargement
                        time.sleep(2)
                        
                        # Trouver tous les messages actuellement visibles
                        message_divs = self.driver.find_elements(By.XPATH, "//div[@role='region']")

                        # Vérifier le nombre de messages
                        current_message_count = len(message_divs)
                        if current_message_count > previous_message_count:
                            previous_message_count = current_message_count
                        else:
                            break  # Sortir si aucun nouveau message n'a été chargé

                        # Faire défiler vers le bas pour charger plus de messages
                        self.driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", message_divs[-1])

                    # Extraire les messages après le défilement
                    for msg_div in message_divs:
                        name, content, timestamp = self.extract_message_info(msg_div)

                        # Générer un identifiant unique pour chaque message
                        message_id = f"{name}_{content}_{timestamp}"

                        # Vérifier si le message est déjà présent
                        if message_id not in self.unique_messages and name and content and timestamp:
                            self.messages.append({
                                "Nom": name,
                                "Message": content,
                                "Heure": timestamp,
                                "Date d'extraction": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            })
                            self.unique_messages.add(message_id)

                    # Retourner à la liste des conversations
                    back_button = WebDriverWait(self.driver, 30).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "[title='Retour']"))
                    )
                    back_button.click()
                except Exception as e:
                    logging.error(f"Erreur lors de l'extraction de la conversation {i}: {e}")
                    continue
        except Exception as e:
            logging.error(f"Erreur lors de la recherche des conversations : {e}")
            raise e

    def read_existing_data(self):
        """
        Lit les données existantes dans le fichier Excel s'il existe.

        Returns:
            list: Liste de dictionnaires contenant les messages existants.
        """
        try:
            existing_data = pd.read_excel('skypeMessages.xlsx')
            return existing_data.to_dict(orient='records')
        except FileNotFoundError:
            logging.info("Aucun fichier Excel existant trouvé, création d'un nouveau fichier.")
            return []  # Aucun fichier trouvé, retourne une liste vide

    def export_to_excel(self):
        """
        Exporte les messages extraits vers un fichier Excel, en ajoutant aux données existantes.
        """
        existing_messages = self.read_existing_data()  # Lire les messages existants

        # Convertir la liste de messages existants en un ensemble pour éviter les doublons
        existing_message_ids = {f"{msg['Nom']}_{msg['Message']}_{msg['Heure']}" for msg in existing_messages}

        # Ajouter les nouveaux messages, en évitant les doublons
        for message in self.messages:
            message_id = f"{message['Nom']}_{message['Message']}_{message['Heure']}"
            if message_id not in existing_message_ids:
                existing_messages.append(message)

        # Créer un DataFrame et l'écrire dans le fichier Excel
        df = pd.DataFrame(existing_messages)
        df.to_excel('skypeMessages.xlsx', index=False)
        logging.info("Messages exportés vers skypeMessages.xlsx.")

    def close(self):
        """
        Ferme le navigateur Selenium.
        """
        logging.info("Fermeture du navigateur.")
        self.driver.quit()


# Utilisation
if __name__ == "__main__":
    extractor = SkypeWebExtractor()

    try:
        extractor.login()  # L'utilisateur saisit ses identifiants directement dans Skype Web.
        extractor.extract_conversations(limit=10)  # Limiter à 10 conversations
        extractor.export_to_excel()  # Exporter dans le fichier skypeMessages.xlsx
    except Exception as e:
        print(f"Erreur: {str(e)}")
    finally:
        extractor.close()
