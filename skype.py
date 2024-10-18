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

                    # Attendre que les messages soient visibles sans utiliser time.sleep()
                    message_divs = WebDriverWait(self.driver, 30).until(
                        EC.presence_of_all_elements_located((By.XPATH, "//div[@role='region']"))
                    )

                    for msg_div in message_divs:
                        name, content, timestamp = self.extract_message_info(msg_div)

                        # Générer un identifiant unique pour chaque message
                        message_id = f"{name}_{content}_{timestamp}"

                        # Vérifier si le message est déjà présent
                        if message_id not in self.unique_messages and name and content and timestamp:
                            self.messages.append({
                                "Nom": name,
                                "Message": content,
                                "Heure": timestamp
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

    def export_to_excel(self, filename=None):
        """
        Exporte les messages extraits dans un fichier Excel.

        Args:
            filename (str, optional): Le nom du fichier Excel. Si aucun nom n'est fourni, un nom par défaut est généré.

        Returns:
            str: Le nom du fichier Excel généré.
        """
        if filename is None:
            filename = f"skype_messages_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        try:
            df = pd.DataFrame(self.messages)
            df.to_excel(filename, index=False)
            logging.info(f"Messages exportés vers {filename}")
            print(f"Messages exportés vers {filename}")
            return filename
        except Exception as e:
            logging.error(f"Erreur lors de l'exportation vers Excel : {e}")
            raise e

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
        extractor.export_to_excel()  # Exporter au format Excel
    except Exception as e:
        print(f"Erreur: {str(e)}")
    finally:
        extractor.close()
