import pandas as pd
import folium
import os

# 📁 Dossier contenant les fichiers Excel
folder_path = r"D:\excel cerales"

# 📌 Liste des fichiers
files = [
    "capacité de stockage minoterie avril 26.xlsx",
    "capacité de stockage oc  avril  26.xlsx",
    "centres de collecte avril 26.xlsx",
    "ecart 2016_2025.xlsx",
    "ecart 2019 et 2025.xlsx",
    "synthèse délégation.xlsx",
    "synthèse gouvernorat.xlsx"
]

# 🗺️ Dictionnaire complet des délégations tunisiennes avec coordonnées GPS
delegations_coords = {
    "Aouina": [36.8667, 10.1833], "Raoued": [36.8833, 10.3667], "Kalâat el-Andalous": [36.8667, 10.2667],
    "Sidi Thabet": [36.9167, 10.3833], "Ben Arous": [36.7667, 10.2333], "Ezzahra": [36.7167, 10.2],
    "Hammam Lif": [36.7333, 10.3667], "Mornag": [36.6833, 10.3667], "Radès": [36.8, 10.25],
    "Bizerte": [37.2833, 9.8667], "Ras Jebel": [37.35, 9.7667], "Sidi El Mekki": [37.1667, 9.9],
    "Raf Raf": [37.3167, 9.9833], "Ghar El Melh": [37.1833, 10.15], "Menzel Bourguiba": [37.1833, 9.8167],
    "Menzel Jemil": [37.2833, 9.8667], "Tinja": [37.2167, 9.95], "Metline": [37.3, 9.75],
    "Béja": [36.7333, 9.1833], "Medjez el-Bab": [36.4, 9.3333], "Testour": [36.5667, 9.45],
    "Teboursouk": [36.45, 9.3667], "Nefza": [36.6167, 8.8333], "Gabès": [33.8833, 10.1],
    "Matmata": [33.7333, 9.75], "Mareth": [33.6333, 9.8333], "Ghannouch": [33.9833, 10.1667],
    "El Hamma": [33.6, 10.5], "Gafsa": [34.4, 8.7833], "Mdhilla": [34.5833, 8.5833],
    "El Guettar": [34.6667, 8.95], "Sened": [34.3, 9.0], "Métlaoui": [34.3667, 8.4167],
    "Redeyef": [34.3333, 8.5667], "Jendouba": [36.5, 8.7667], "Bou Arada": [36.3667, 8.9833],
    "Goubellat": [36.2833, 9.1167], "Tabarka": [36.9667, 8.75], "Aïn Drahem": [36.8333, 8.65],
    "Kairouan": [35.6667, 9.85], "Oueslatia": [35.6667, 9.65], "Sbeitla": [35.7667, 9.2667],
    "Haffouz": [35.5333, 9.7667], "Nasrallah": [35.6, 9.95], "Sidi Saïd": [35.7167, 9.75],
    "Kasserine": [35.1667, 8.8333], "Foussana": [35.3167, 8.5667], "Thala": [35.4333, 8.6167],
    "Feriana": [35.5167, 8.9167], "Kébili": [33.7333, 8.9667], "Douz": [33.4667, 8.8167],
    "Souk Lahad": [33.7333, 8.7667], "Jemna": [33.5, 8.9167], "Mahdia": [35.5, 11.0667],
    "Chebba": [35.2333, 11.1167], "El Jem": [35.3, 11.1], "Skhira": [34.7167, 11.0667],
    "Manouba": [36.8, 10.1], "Mornaguia": [36.8333, 10.0667], "Tébourba": [36.7667, 9.95],
    "Médenine": [33.3667, 10.5], "Djerba": [33.8, 10.85], "Zarzis": [33.5, 11.1167],
    "Ben Guerdane": [33.2, 10.95], "Ajim": [33.75, 10.85], "Monastir": [35.7667, 10.8333],
    "Jemmal": [35.7167, 10.7667], "Sahline": [35.7667, 10.9167], "Teboulba": [35.8, 10.85],
    "Bekalta": [35.8333, 10.9833], "Nabeul": [36.4333, 10.7333], "Hammamet": [36.4, 10.6167],
    "Kelibia": [36.8333, 11.1], "Korba": [36.5667, 11.0167], "Menzel Temime": [36.7333, 10.95],
    "Dar Chaabane": [36.3833, 10.5667], "Sfax": [34.7333, 10.75], "Sakiet El Daiesa": [34.6, 10.9167],
    "Agareb": [34.8333, 11.1], "El Amra": [34.8667, 10.85], "Kerkennah": [34.6, 11.2667],
    "Thyna": [34.5833, 10.7667], "Sidi Bouzid": [35.0333, 9.4833], "Regueb": [35.2667, 9.2667],
    "Mèknès": [34.8333, 9.4167], "Bir El Hafey": [35.0333, 9.5667], "Sousse": [35.8333, 10.6333],
    "Akouda": [35.8667, 10.65], "Msaken": [35.7667, 10.55], "Kalâa Kebira": [35.8667, 10.5],
    "Sidi El Hani": [35.95, 10.5667], "Hammam Sousse": [35.85, 10.6167], "Bouficha": [35.9, 10.7],
    "Tataouine": [32.9333, 10.45], "Remada": [32.8333, 9.95], "Dehiba": [32.85, 10.2167],
    "Ghomrassen": [33.2167, 10.1167], "Foum Tataouine": [32.95, 10.3667], "Tozeur": [33.9167, 8.1333],
    "Nefta": [33.8667, 7.8667], "Tunis": [36.8, 10.1833], "Carthage": [36.8667, 10.3167],
    "La Marsa": [36.8667, 10.3333], "La Goulette": [36.8, 10.3167], "Le Kef": [36.1667, 8.7167],
    "Siliana": [36.0833, 9.35],
}

# 🌍 Carte centrée sur la Tunisie
m = folium.Map(location=[34.0, 9.0], zoom_start=6)

# 🔁 Boucle sur chaque fichier
for file in files:
    file_path = os.path.join(folder_path, file)

    try:
        # 📖 Lecture intelligente (Excel ou CSV)
        if file == "ecart 2016_2025.xlsx":
            try:
                df = pd.read_csv(file_path, sep=',')
            except:
                df = pd.read_csv(file_path, sep=';')
        else:
            df = pd.read_excel(file_path)

        print(f"\n📄 {file}")
        print("Colonnes :", df.columns.tolist())

        # 🔍 Détection automatique des colonnes Latitude / Longitude
        lat_col = None
        lon_col = None
        delegation_col = None

        for col in df.columns:
            col_clean = col.lower().strip()

            if "lat" in col_clean:
                lat_col = col
            if "lon" in col_clean:
                lon_col = col
            if "delegation" in col_clean or "délégation" in col_clean:
                delegation_col = col

        # 📌 Groupe (checkbox)
        fg = folium.FeatureGroup(name=file)

        for _, row in df.iterrows():
            lat = None
            lon = None

            # 🔢 Si colonnes lat/lon existent
            if lat_col and lon_col:
                lat = pd.to_numeric(row[lat_col], errors='coerce')
                lon = pd.to_numeric(row[lon_col], errors='coerce')

            # 🗺️ Sinon, chercher par délégation
            elif delegation_col:
                delegation = str(row[delegation_col]).strip()
                if delegation in delegations_coords:
                    lat, lon = delegations_coords[delegation]

            # ❌ Si pas de coordonnées
            if lat is None or lon is None or pd.isna(lat) or pd.isna(lon):
                continue

            # 🧾 Popup avec tous les champs
            popup_text = "<div style='font-size:12px'>"
            for col in df.columns:
                popup_text += f"<b>{col}:</b> {row[col]}<br>"
            popup_text += "</div>"

            folium.Marker(
                location=[lat, lon],
                popup=folium.Popup(popup_text, max_width=300),
                icon=folium.Icon(color="blue", icon="info-sign")
            ).add_to(fg)

        fg.add_to(m)

    except Exception as e:
        print(f"❌ Erreur avec {file}: {e}")

# 🎛️ Contrôle des couches
folium.LayerControl(collapsed=False).add_to(m)

# 💾 Sauvegarde
output_path = r"D:\excel cerales\index.html"
m.save(output_path)

print(f"\n✅ Carte générée avec succès : {output_path}")
