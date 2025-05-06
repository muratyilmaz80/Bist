from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Safari'yi başlat
driver = webdriver.Safari()

# Teslanın ana sayfasını aç
url = "https://www.tesla.com/tr_tr/inventory/new/my"
driver.get(url)

# Çerez onayı sayfasını bekle ve onayla/kapat
try:
    # Çerez onayı butonunu bekle ve tıkla
    wait = WebDriverWait(driver, 20)  # 20 saniyelik daha uzun bir bekleme süresi
    accept_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-test='accept-cookies']")))

    # Çerez onayı butonunu tıklama
    accept_button.click()
    print("✅ Çerez onayı kabul edildi.")

except Exception as e:
    print("⚠️ Çerez onayı bulunamadı veya hata oluştu:", str(e))

# Sayfanın yüklenmesini bekle
try:
    wait.until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR, "div[data-test='vehicleCard']")))
    print("✅ Araç kartları yüklendi.")
except Exception as e:
    print("⚠️ Araç kartları yüklenmedi veya hata oluştu:", str(e))

# Araç kartlarını al
vehicles = driver.find_elements(By.CSS_SELECTOR, "div[data-test='vehicleCard']")

# Araç bilgilerini yazdır
if not vehicles:
    print("❌ Sayfada araç bulunamadı.")
else:
    print(f"🚗 Toplam {len(vehicles)} araç bulundu.\n")
    for idx, v in enumerate(vehicles, start=1):
        print(f"🔹 Araç #{idx}")
        print(v.text)
        print("-" * 50)

# Tarayıcıyı kapat
driver.quit()
