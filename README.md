# 🔍 CheckVersions

Un outil de bureau Windows permettant de vérifier si vos logiciels installés sont à jour, en comparant la version locale avec la dernière version disponible en ligne.

---

## 📋 Fonctionnalités

- Interface graphique simple avec cases à cocher par logiciel
- Vérification simultanée de plusieurs applications
- Comparaison automatique version locale / version en ligne
- Affichage coloré des résultats :
  - ✅ **Vert** — à jour
  - ❌ **Rouge** — mise à jour disponible
  - ⚠️ **Jaune** — non installé ou version introuvable
- Bouton **"All"** pour tout sélectionner/désélectionner
- Exécution en arrière-plan (sans bloquer l'interface)

---

## 🖥️ Logiciels supportés

| Catégorie | Logiciels |
|---|---|
| Navigateurs | Chrome, Firefox, Microsoft Edge |
| Bureau à distance | RustDesk, AnyDesk |
| Cloud | NextCloud |
| Antivirus | Malwarebytes, ESET Endpoint Security, ESET Full Disk Encryption |
| Bureautique | OnlyOffice, Adobe Acrobat Reader |
| Nettoyage | BleachBit, Dell Support Assist |
| Sauvegarde | Veeam Agent for Microsoft Windows |
| Utilitaires | 7-Zip, VLC |
| VPN | OpenVPN Connect |

---

## ⚙️ Prérequis

- **OS** : Windows 10 / 11
- **Python** : 3.9+

### Dépendances Python
```bash
pip install requests customtkinter packaging colorama beautifulsoup4 feedparser cloudscraper selenium pywin32
```

---

## 🚀 Utilisation

### Lancer depuis Python
```bash
python CheckVersions.py
```

### Compiler en `.exe`
```bash
pyinstaller --onefile --collect-all selenium CheckVersions.py
```

L'exécutable sera généré dans le dossier `dist/`.

---

## 📸 Aperçu
```
✅ | Chrome          | local : v136.0.7103 | online : v136.0.7103
❌ | Firefox         | local : v128.0.0    | online : v139.0.1
⚠️ | RustDesk        | local : non installé ou version introuvable
```

---

## 🗂️ Structure du projet
```
CheckVersions/
├── CheckVersions.exe
├── CheckVersions.py   # Script principal
└── README.md
```

---

## ⚠️ Notes

- **Dell Support Assist** est détecté via le setup (`SupportAssistinstaller.exe`) et non via le programme installé — assurez-vous que le fichier d'installation est présent dans vos dossiers habituels (Téléchargements, Bureau, etc.).
- **Malwarebytes** est détecté via le registre Windows.
- **Veeam** nécessite Chrome installé sur la machine pour le scraping via Selenium.
- Certains sites utilisent une protection anti-bot (Cloudflare) — `cloudscraper` est utilisé pour les contourner.

---

## 📄 Licence

Projet personnel — libre d'utilisation et de modification.