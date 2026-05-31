
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

def main():
    print("Oturum açılıyor...")
    tgt = tgt_al(KULLANICI_ADI, SIFRE)
    if not tgt:
        print("Giriş başarısız.")
        return

    headers = {
        "TGT": tgt,
        "Content-Type": "application/json",
        "Accept": "application/json"
    }

    # 1. Dağıtım Şirketlerini Çek
    print("Dağıtım Şirketleri listesi alınıyor...")
    dists = fetch_json_post(DIST_URL, {"period": DONEM}, headers)
    print(f"Toplam {len(dists)} dağıtım şirketi bulundu.")

    all_data = []

    # 2. Her Dağıtım Şirketi İçin Döngü
    for dist in dists:
        dist_id = dist.get('id')
        dist_name = dist.get('name')
        
        print(f"\nİşleniyor: {dist_name} ({dist_id})")
        
        # 3. Dağıtım Şirketine Ait Profil Gruplarını Çek
        profiles = fetch_json_post(PROFILE_URL, {"period": DONEM, "distributionId": dist_id}, headers)
        print(f"  -> {len(profiles)} profil grubu bulundu.")
        
        for prof in profiles:
            prof_id = prof.get('id')
            prof_name = prof.get('name')
            
            # 4. Sayaç Okuma Tipleri (1 ve 3) İçin Döngü
            for meter_type in [1, 3]:
                # Veri İsteği
                req_body = {
                    "period": DONEM,
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
                                    "DONEM": DONEM,
                                    "Organizasyon ID": dist_id,
                                    "Organizasyon Adı": dist_name,
                                    "Profil Grup ID": prof_id,
                                    "Profil Grup Adı": prof_name,
                                    "Sayaç Tipi": meter_type,
                                    "Zaman": item.get("time"),
                                    "Çarpan": item.get("multiplier")
                                })
                            # print(f"    -> Profil: {prof_name} | Tip: {meter_type} | Veri: {len(items)} adet")
                        else:
                             pass 
                             # print(f"    -> Profil: {prof_name} | Tip: {meter_type} | Veri YOK")
                    else:
                        print(f"    -> Hata: {res.status_code} - {res.text}")
                except Exception as e:
                    print(f"    -> İstek Hatası: {e}")
                
                # API sunucusunu boğmamak için kısa bekleme
                # time.sleep(0.05)

    # 5. Excel'e Kaydet
    if all_data:
        print(f"\nToplam {len(all_data)} satır veri toplandı.")
        print("Excel dosyası oluşturuluyor...")
        df = pd.DataFrame(all_data)
        try:
            df.to_excel(OUTPUT_FILE, index=False)
            print(f"Tamamlandı: {OUTPUT_FILE}")
        except Exception as e:
            print(f"Excel kaydetme hatası: {e}")
    else:
        print("\nHiç veri toplanamadı.")

if __name__ == "__main__":
    main()

