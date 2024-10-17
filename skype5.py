from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
from datetime import datetime

class SkypeWebExtractor:
    def __init__(self):
        self.driver = webdriver.Chrome()
        self.messages = []
        
    def login(self, email, password):
        print("Connexion à Skype Web...")
        self.driver.get("https://web.skype.com")
        time.sleep(5)
        
        # Login email
        email_input = WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.NAME, "loginfmt"))
        )
        email_input.send_keys(email)
        self.driver.find_element(By.ID, "idSIButton9").click()
        
        # Login password
        password_input = WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.NAME, "passwd"))
        )
        password_input.send_keys(password)
        self.driver.find_element(By.ID, "idSIButton9").click()
        
        print("Attente du chargement de Skype...")
        input("Attendez le chargement des messages et appuyez sur Entrée pour continuer...")
        
    def extract_message_info(self, message_div):
        aria_label = message_div.get_attribute("aria-label")
        parts = aria_label.split(", ")
        
        if len(parts) >= 3:
            name = parts[0]
            message = ", ".join(parts[1:-1])  # Le message peut contenir des virgules
            time = parts[-1].split(" à ")[-1]
            return name, message, time
        return None, None, None

    def extract_conversations(self):
        print("Recherche des conversations...")
        time.sleep(5)
        
        conversation_list = WebDriverWait(self.driver, 30).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "[role='listitem']"))
        )
        
        print(f"Nombre de conversations trouvées : {len(conversation_list)}")
        
        for i, conv in enumerate(conversation_list, 1):
            try:
                print(f"Extraction conversation {i}/{len(conversation_list)}")
                
                conv.click()
                time.sleep(5)
                
                # Extraire les messages
                message_divs = WebDriverWait(self.driver, 30).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//div[@role='region']"))
                )
                
                for msg_div in message_divs:
                    name, content, timestamp = self.extract_message_info(msg_div)
                    if name and content and timestamp:
                        self.messages.append({
                            "Nom": name,
                            "Message": content,
                            "Heure": timestamp
                        })
                
                # Retourner à la liste des conversations
                back_button = self.driver.find_element(By.CSS_SELECTOR, "[title='Retour']")
                back_button.click()
                time.sleep(3)
                
            except Exception as e:
                print(f"Erreur sur la conversation {i}: {str(e)}")
                continue
    
    def export_to_csv(self, filename=None):
        if filename is None:
            filename = f"skype_messages_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            
        df = pd.DataFrame(self.messages)
        df.to_csv(filename, index=False, encoding='utf-8')
        print(f"Messages exportés vers {filename}")
        return filename
        
    def close(self):
        self.driver.quit()

# Utilisation
if __name__ == "__main__":
    extractor = SkypeWebExtractor()
    
    try:
        extractor.login("yaoandersonkouassi31@gmail.com", "Mon_compte_skype")
        time.sleep(5)  # Attendre que tout soit bien chargé
        
        extractor.extract_conversations()
        extractor.export_to_csv()
        
    except Exception as e:
        print(f"Erreur: {str(e)}")
    finally:
        extractor.close()