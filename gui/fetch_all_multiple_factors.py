
import requests
import pandas as pd
import json
import time

# Yapılandırma
KULLANICI_ADI = ""
SIFRE = ""
AUTH_URL = "https://giris.epias.com.tr/cas/v1/tickets"
DIST_URL = "https://seffaflik.epias.com.tr/electricity-service/v1/consumption/data/multiple-factor-distribution"
PROFILE_URL = "https://seffaflik.epias.com.tr/electricity-service/v1/consumption/data/multiple-factor-profile-group"
DATA_URL = "https://seffaflik.epias.com.tr/electricity-service/v1/consumption/data/multiple-factor"

OUTPUT_FILE = "Tum_Profil_Katsayilari.xlsx"
DONEM = "2025-12-01T00:00:00+03:00"

def tgt_al(kullanici_adi, sifre):
    payload = {"username": kullanici_adi, "password": sifre}
    try:
        yanit = requests.post(AUTH_URL, data=payload)
        yanit.raise_for_status()
        tgt_location = yanit.headers.get("Location", "")
        if "TGT" not in tgt_location:
            return None
        return tgt_location.split("/")[-1].strip()
    except Exception as e:
        print(f"Auth Error: {e}")
        return None

def fetch_json_post(url, body, headers):
    try:
        res = requests.post(url, json=body, headers=headers)
        if res.status_code == 200:
            data = res.json()
            if isinstance(data, dict) and 'items' in data:
                return data['items']
            elif isinstance(data, list):
                return data
        return []
    except:
        return []

def run_process(donem_input, output_folder, username, password, progress_callback=None):
    """
    GUI veya dışarıdan çağrılabilmesi için ana işlem fonksiyonu.
    progress_callback: f(percent, message) formatında fonksiyon.
    """
    if progress_callback: progress_callback(0, "Hazırlanıyor...")

    # Tarih formatı kontrol ve dönüşüm
    # Gelen format: YYYY-MM-DD -> Hedef: YYYY-MM-DDT00:00:00+03:00
    if len(donem_input) == 10 and "-" in donem_input:
        donem_iso = f"{donem_input}T00:00:00+03:00"
    else:
        donem_iso = donem_input 

    # Çıktı dosya yolu
    full_output_path = f"{output_folder}/{OUTPUT_FILE}" if output_folder else OUTPUT_FILE

    print(f"İşlem Başlıyor: {donem_iso}")
    
    if progress_callback: progress_callback(0.1, "Oturum açılıyor...")
    print("Oturum açılıyor...")
    tgt = tgt_al(username, password)
    if not tgt:
        return False, "Giriş başarısız (Kullanıcı adı/Şifre hatalı veya TGT alınamadı)."

    headers = {
        "TGT": tgt,
        "Content-Type": "application/json",
        "Accept": "application/json"
    }

    # 1. Dağıtım Şirketlerini Çek
    if progress_callback: progress_callback(0.2, "Dağıtım Şirketleri listeleniyor...")
    print("Dağıtım Şirketleri listesi alınıyor...")
    dists = fetch_json_post(DIST_URL, {"period": donem_iso}, headers)
    print(f"Toplam {len(dists)} dağıtım şirketi bulundu.")

    all_data = []

    # 2. Her Dağıtım Şirketi İçin Döngü
    total_dists = len(dists)
    for i, dist in enumerate(dists):
        # Progress Calculation: 0.2 to 0.9 range
        if progress_callback:
            percent = 0.2 + (0.7 * (i / total_dists))
            progress_callback(percent, f"İşleniyor: {dist.get('name')}")

        dist_id = dist.get('id')
        dist_name = dist.get('name')
        
        print(f"\nİşleniyor: {dist_name} ({dist_id})")
        
        # 3. Dağıtım Şirketine Ait Profil Gruplarını Çek
        profiles = fetch_json_post(PROFILE_URL, {"period": donem_iso, "distributionId": dist_id}, headers)
        print(f"  -> {len(profiles)} profil grubu bulundu.")
        
        for prof in profiles:
            prof_id = prof.get('id')
            prof_name = prof.get('name')
            
            # 4. Sayaç Okuma Tipleri (1 ve 3) İçin Döngü
            for meter_type in [1, 3]:
                # Veri İsteği
                req_body = {
                    "period": donem_iso,
                    "distributionId": dist_id,
                    "meterReadingType": meter_type,
                    "subscriberProfileGroup": prof_id,
                    "page": {
                        "number": 1,
                        "size": 1000, # Ayın tüm verisini almak için büyük bir sayı
                        "sort": {"direction": "ASC", "field": "date"}
                    }
                }
                
                try:
                    res = requests.post(DATA_URL, json=req_body, headers=headers)
                    if res.status_code == 200:
                        resp_json = res.json()
                        items = resp_json.get("items", [])
                        
                        if items:
                            # Verileri listeye ekle
                            for item in items:
                                all_data.append({
                                    "DONEM": donem_iso,
                                    "Organizasyon ID": dist_id,
                                    "Organizasyon Adı": dist_name,
                                    "Profil Grup ID": prof_id,
                                    "Profil Grup Adı": prof_name,
                                    "Sayaç Tipi": meter_type,
                                    "Zaman": item.get("time"),
                                    "Çarpan": item.get("multiplier")
                                })
                    else:
                        print(f"    -> Hata: {res.status_code} - {res.text}")
                except Exception as e:
                    print(f"    -> İstek Hatası: {e}")

    # 5. Excel'e Kaydet
    if all_data:
        if progress_callback: progress_callback(0.95, "Excel oluşturuluyor...")
        print(f"\nToplam {len(all_data)} satır veri toplandı.")
        print("Excel dosyası oluşturuluyor...")
        df = pd.DataFrame(all_data)
        try:
            df.to_excel(full_output_path, index=False)
            print(f"Tamamlandı: {full_output_path}")
            if progress_callback: progress_callback(1.0, "Tamamlandı.")
            return True, f"Tamamlandı. {len(all_data)} satır veri indirildi.\nDosya: {full_output_path}"
        except Exception as e:
            error_msg = f"Excel kaydetme hatası: {e}"
            print(error_msg)
            if progress_callback: progress_callback(0, "Hata oluştu.")
            return False, error_msg
    else:
        print("\nHiç veri toplanamadı.")
        if progress_callback: progress_callback(0, "Veri yok.")
        return False, "Hiç veri toplanamadı."

def main():
    # Varsayılan çalışma şekli (Dosya direkt çalıştırıldığında)
    run_process(DONEM, "", KULLANICI_ADI, SIFRE)

if __name__ == "__main__":
    main()
