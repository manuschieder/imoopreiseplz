import openpyxl
from bs4 import BeautifulSoup
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

# URL der Zielwebsite
url = "https://www.interhyp.de/immobilienpreise/#immobilienpreisentwicklung"

# Selenium WebDriver einrichten
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 15)

# Excel-Datei einrichten
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Immobilienpreise"
sheet.append(["PLZ", "Marktwert Haus 2015", "Marktwert Haus Aktuell", "Marktwert Wohnung 2015", "Marktwert Wohnung Aktuell"])

cookie_banner_closed = False

# Funktion zum Loggen mit Zeitstempel
def log_with_timestamp(message):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}")

# PLZ aus Excel-Datei einlesen
def read_plz_from_excel(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    plz_list = [str(row[0].value).zfill(5) for row in sheet.iter_rows(min_row=2, max_col=1)]
    return plz_list

def scrape_data_selenium(plz):
    """Daten für eine einzelne PLZ mit Selenium abfragen."""
    global cookie_banner_closed

    driver.get(url)
    time.sleep(5)  # Warten, bis die Seite vollständig geladen ist

    if not cookie_banner_closed:
        try:
            cookie_button = wait.until(EC.element_to_be_clickable((By.ID, "cm-btnAcceptAll")))
            cookie_button.click()
            log_with_timestamp("Cookie-Banner geschlossen.")
            cookie_banner_closed = True
        except Exception as e:
            log_with_timestamp("Kein Cookie-Banner gefunden oder Fehler beim Schließen: " + str(e))

    try:
        # Eingabe der PLZ
        input_field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='Postleitzahl / Ort']")))
        input_field.clear()
        input_field.send_keys(Keys.CONTROL + "a")  # Sicherstellen, dass alles gelöscht wird
        input_field.send_keys(Keys.DELETE)  # Löschen der Inhalte
        input_field.send_keys(plz)
        input_field.send_keys(Keys.RETURN)

        # Wartezeit für die Aktualisierung der Daten
        time.sleep(8)

        # Extrahiere die Werte für Marktwert Haus
        try:
            haus_div = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "content_LXx3p4L4YM")))[0]
            haus_soup = BeautifulSoup(haus_div.get_attribute('innerHTML'), 'html.parser')
            marktwert_2015 = haus_soup.select_one("p:contains('2015:') strong").get_text(strip=True)
            marktwert_aktuell = haus_soup.select_one("p:contains('Aktuell:') strong").get_text(strip=True)
        except Exception as e:
            marktwert_2015 = marktwert_aktuell = "N/A"
            log_with_timestamp(f"Fehler beim Auslesen der Marktwerte Haus für PLZ {plz}: {e}")

        # Extrahiere die Werte für Marktwert Wohnung
        try:
            wohnung_div = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "content_LXx3p4L4YM")))[1]  # Zweites Div für Wohnung
            wohnung_soup = BeautifulSoup(wohnung_div.get_attribute('innerHTML'), 'html.parser')
            marktwert_wohnung_2015 = wohnung_soup.select_one("p:contains('2015:') strong").get_text(strip=True)
            marktwert_wohnung_aktuell = wohnung_soup.select_one("p:contains('Aktuell:') strong").get_text(strip=True)
        except Exception as e:
            marktwert_wohnung_2015 = marktwert_wohnung_aktuell = "N/A"
            log_with_timestamp(f"Fehler beim Auslesen der Marktwerte Wohnung für PLZ {plz}: {e}")

        # Daten in Excel schreiben
        sheet.append([plz, marktwert_2015, marktwert_aktuell, marktwert_wohnung_2015, marktwert_wohnung_aktuell])

    except Exception as e:
        log_with_timestamp(f"Fehler bei PLZ {plz}: {e}")

# Pfad zur Excel-Datei mit PLZ
input_excel_file = "PLZ_Liste.xlsx"

# PLZ-Liste aus Excel einlesen
deutsche_plz = read_plz_from_excel(input_excel_file)

# Alle PLZ durchlaufen
for plz in deutsche_plz:
    try:
        log_with_timestamp(f"Abfrage für PLZ {plz}...")
        scrape_data_selenium(plz)
    except Exception as e:
        log_with_timestamp(f"Fehler bei PLZ {plz}: {e}")

# Selenium WebDriver schließen
driver.quit()

# Excel-Datei speichern
excel_file = "Immobilienpreise.xlsx"
workbook.save(excel_file)
log_with_timestamp(f"Daten wurden erfolgreich in '{excel_file}' gespeichert.")
