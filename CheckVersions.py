import requests
import customtkinter as ctk
import os
import subprocess
import re
from packaging.version import Version, InvalidVersion
from colorama import *
from bs4 import BeautifulSoup
import feedparser
import cloudscraper
import sys
import argparse
import ctypes
from ctypes import wintypes
import win32api
import winreg
import threading
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


"""
IMPORTANT POUR COMPILER : pyinstaller --onefile --collect-all selenium CheckVersions.py
"""

app = ctk.CTk()
app.geometry("575x400")
app.title("Verification des versions")

frame_options = ctk.CTkFrame(app)
frame_options.grid(pady=10)

cols = 3

options = ["Chrome", "Firefox", "Microsoft Edge", "RustDesk", "NextCloud", "Malwarebytes", "ESET Endpoint Security", 
           "ESET Full Disk Encryption for Windows", "OnlyOffice", "BleachBit", "Adobe Acrobat Reader", "AnyDesk", "Dell Support Assist", 
           "Veeam Agent For Microsoft Windows", "7-Zip", "Open VPN Connect",]
vars = []

#Navigateur
installchrome = r"(Get-Item 'C:\Program Files\Google\Chrome\Application\chrome.exe').VersionInfo.ProductVersion"
urlchrome = "https://chromereleases.googleblog.com/feeds/posts/default?alt=rss"
installfirefox = r"(Get-Item 'C:\Program Files\Mozilla Firefox\firefox.exe').VersionInfo.ProductVersion"
urlfirefox = "https://www.firefox.com/en-US/releases/"
installmsgedge = r"(Get-Item 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe').VersionInfo.ProductVersion"
urlmsedge = "https://learn.microsoft.com/en-us/deployedge/microsoft-edge-relnote-stable-channel"

#Remote Desktop
installrustdesk = r"(Get-Item 'C:\Program Files\RustDesk\rustdesk.exe').VersionInfo.ProductVersion"
urlrustdesk = "https://api.github.com/repos/rustdesk/rustdesk/releases/latest"
installanydesk = r'(Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-Object { $_.DisplayName -like "*AnyDesk*" } | Select-Object -ExpandProperty DisplayVersion) -replace "ad ",""'
urlanydesk = "https://anydesk.com/fr/changelog/windows"

#Nextcloud
installnextcloud = r"(Get-Item 'C:\Program Files\Nextcloud\nextcloud.exe').VersionInfo.ProductVersion"
urlnextcloud = "https://api.github.com/repos/nextcloud/desktop/releases/latest"

#Anti Virus
installmalwarebytes = r"(Get-Item 'C:\Program Files\Malwarebytes\Anti-Malware\Malwarebytes.exe').VersionInfo.ProductVersion"
urlmalwarebytes = "https://help.malwarebytes.com/api/v2/help_center/en-us/articles"
installesetendpointsecurity = r"(Get-Item 'C:\Program Files\ESET\ESET Security\ecmds.exe').VersionInfo.ProductVersion"
installesetfulldiskencryption = r"(Get-Item 'C:\Program Files\ESET\ESET Full Disk Encryption\EFDEUI.exe').VersionInfo.ProductVersion"
urleset = "https://help.eset.com/latestVersions/?lang=en-US"
#Offices
installonlyoffice = r"(Get-Item 'C:\Program Files\ONLYOFFICE\DesktopEditors\DesktopEditors.exe').VersionInfo.ProductVersion"
urlonlyoffice = "https://api.github.com/repos/ONLYOFFICE/DesktopEditors/releases/latest"
installaar = r"(Get-Item 'C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe').VersionInfo.ProductVersion"
urlaar = "https://www.adobe.com/devnet-docs/acrobatetk/tools/ReleaseNotesDC/index.html"

#Cleaner
installbleachbit = r"(Get-Item 'C:\Program Files (x86)\BleachBit\bleachbit.exe').VersionInfo.ProductVersion"
urlbleachbit = "https://api.github.com/repos/bleachbit/bleachbit/releases/latest"
installdsp = r"(Get-Item 'C:\Program Files\Dell\SupportAssistAgent\bin\SupportAssist.exe').VersionInfo.ProductVersion"

#Backup
installveeam = r"(Get-Item 'C:\Program Files\Veeam\Endpoint Backup\Veeam.EndPoint.Tray.exe').VersionInfo.ProductVersion"
urlveeam = "https://www.veeam.com/products/downloads/latest-version.html"

#File Manager
install7zip = r"(Get-Item 'C:\Program Files\7-Zip\7zFM.exe').VersionInfo.ProductVersion"
url7zip = "https://api.github.com/repos/ip7z/7zip/releases/latest"

#VPN
installopenvpn = r"(Get-Item 'C:\Program Files\OpenVPN Connect\OpenVPNConnect.exe').VersionInfo.ProductVersion"
urlopenvpn = "https://openvpn.net/connect-docs/windows-release-notes.html"

url_commands = {
    "BleachBit": urlbleachbit,
    "OnlyOffice": urlonlyoffice,
    "RustDesk": urlrustdesk,
    "NextCloud": urlnextcloud,
    "Firefox": urlfirefox,
    "Adobe Acrobat Reader": urlaar,
    "Chrome": urlchrome,
    "AnyDesk": urlanydesk,
    "Malwarebytes": urlmalwarebytes,
    "ESET Endpoint Security": urleset,
    "ESET Full Disk Encryption for Windows": urleset,
    "Microsoft Edge": urlmsedge,
    "Veeam Agent For Microsoft Windows": urlveeam,
    "7-Zip": url7zip,
    "Open VPN Connect": urlopenvpn,
}

install_commands = {
    "Chrome": installchrome,
    "Firefox": installfirefox,
    "Microsoft Edge": installmsgedge,
    "RustDesk": installrustdesk,
    "AnyDesk": installanydesk,
    "NextCloud": installnextcloud,
    "Malwarebytes": installmalwarebytes,
    "ESET Endpoint Security": installesetendpointsecurity,
    "ESET Full Disk Encryption for Windows": installesetfulldiskencryption,
    "OnlyOffice": installonlyoffice,
    "BleachBit": installbleachbit,
    "Adobe Acrobat Reader": installaar,
    "Veeam Agent For Microsoft Windows": installveeam,
    "7-Zip": install7zip,
    "Open VPN Connect": installopenvpn,
}

cols = 3

def get_malwarebytes_version():
    keys = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]
    
    for key_path in keys:
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path)
            for i in range(winreg.QueryInfoKey(key)[0]):
                try:
                    subkey_name = winreg.EnumKey(key, i)
                    subkey = winreg.OpenKey(key, subkey_name)
                    try:
                        name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                        if "malwarebytes" in name.lower():
                            version = winreg.QueryValueEx(subkey, "DisplayVersion")[0]
                            return version
                    except FileNotFoundError:
                        pass
                except OSError:
                    pass
        except OSError:
            pass
    
    return "Malwarebytes non trouvé"

def get_dell_support_assist():

    DELL_URL = ("https://www.dell.com/support/manuals/fr-fr/"
                "dell-supportassist-pcs-tablets/sahomepcs_rn/"
                "version-and-validity?guid=guid-b60e62ae-7334-41d6-941b-a60fd380c80f&lang=en-us")

    def find_installer(name="SupportAssistinstaller.exe", whole=False):
        candidates = []
        common = [
            os.path.join(os.environ.get("USERPROFILE",""), "Downloads"),
            os.path.join(os.environ.get("USERPROFILE",""), "Desktop"),
            os.path.join(os.environ.get("USERPROFILE",""), "Documents"),
            os.environ.get("ProgramFiles"),
            os.environ.get("ProgramFiles(x86)"),
            os.environ.get("TEMP"),
            os.environ.get("LOCALAPPDATA"),
        ]
        for base in filter(None, common):
            for root, _, files in os.walk(base):
                if name in files:
                    candidates.append(os.path.join(root, name))
        if candidates or not whole:
            return list(dict.fromkeys(candidates))
        for root, _, files in os.walk("C:\\"):
            if name in files:
                candidates.append(os.path.join(root, name))
        return list(dict.fromkeys(candidates))

    def file_version_pywin(path):
        try:
            info = win32api.GetFileVersionInfo(path, "\\")
            trans = win32api.VerQueryValue(info, r"\\VarFileInfo\\Translation")
            if trans and len(trans) > 0:
                lang, codepage = trans[0]
                key = r"\\StringFileInfo\\%04x%04x\\FileVersion" % (lang, codepage)
                ver = win32api.VerQueryValue(info, key)
                if ver: return ver.strip()
            ver = win32api.VerQueryValue(info, r"\\StringFileInfo\\040904b0\\FileVersion")
            if ver: return ver.strip()
        except Exception:
            return None
        return None

    def file_version_ctypes(path):
        try:
            GetFileVersionInfoSizeW = ctypes.windll.version.GetFileVersionInfoSizeW
            GetFileVersionInfoW = ctypes.windll.version.GetFileVersionInfoW
            VerQueryValueW = ctypes.windll.version.VerQueryValueW
            size = GetFileVersionInfoSizeW(path, None)
            if not size: return None
            res = ctypes.create_string_buffer(size)
            if not GetFileVersionInfoW(path, 0, size, res): return None
            lptr = ctypes.c_void_p()
            lsize = wintypes.UINT()
            if VerQueryValueW(res, "\\VarFileInfo\\Translation", ctypes.byref(lptr), ctypes.byref(lsize)):
                arr_type = ctypes.c_ushort * (lsize.value // ctypes.sizeof(ctypes.c_ushort))
                arr = ctypes.cast(lptr, ctypes.POINTER(arr_type)).contents
                if len(arr) >= 2:
                    lang = arr[0]; codepage = arr[1]
                    sub = "\\StringFileInfo\\%04x%04x\\FileVersion" % (lang, codepage)
                    if VerQueryValueW(res, sub, ctypes.byref(lptr), ctypes.byref(lsize)):
                        ptr = ctypes.cast(lptr, ctypes.c_wchar_p)
                        return ptr.value.strip()
            if VerQueryValueW(res, "\\StringFileInfo\\040904b0\\FileVersion", ctypes.byref(lptr), ctypes.byref(lsize)):
                ptr = ctypes.cast(lptr, ctypes.c_wchar_p)
                return ptr.value.strip()
        except Exception:
            return None
        return None

    def get_file_version(path):
        if not os.path.isfile(path): return None
        if win32api:
            v = file_version_pywin(path)
            if v: return v
        return file_version_ctypes(path)

    def fetch_latest_from_dell(url=DELL_URL):
        r = requests.get(url, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        text = soup.get_text(" ")
        found = re.findall(r'\b\d+(?:\.\d+){1,4}\b', text)
        uniq = list(dict.fromkeys(found))
        def keyfn(s): return [int(x) for x in re.findall(r'\d+', s)]
        uniq.sort(key=keyfn, reverse=True)
        return uniq[0] if uniq else None

    def normalize(v, keep_build=False):
        if not v: return None
        parts = re.findall(r'\d+', v)
        if not parts: return None
        if keep_build: return ".".join(parts)
        if len(parts) > 3: return ".".join(parts[:-1])
        return ".".join(parts)

    def cmp_versions(a, b):
        if a is None or b is None: return None
        A = [int(x) for x in a.split('.')]
        B = [int(x) for x in b.split('.')]
        L = max(len(A), len(B)); A += [0]*(L-len(A)); B += [0]*(L-len(B))
        if A == B: return 0
        return 1 if A > B else -1

    def main():
        parser = argparse.ArgumentParser()
        parser.add_argument("--path", "-p", help="Chemin direct vers le setup (skip search)")
        parser.add_argument("--all", "-a", action="store_true", help="Chercher tout le disque C:")
        parser.add_argument("--keep-build", action="store_true", help="Ne pas retirer le build")
        args = parser.parse_args()

        if args.path:
            installers = [args.path] if os.path.isfile(args.path) else []
        else:
            installers = find_installer(whole=args.all)

        if not installers:
            return ("⚠️ | Dell Support Assist | Application non installé / Setup supprimé")
            sys.exit(1)

        chosen = installers[0]
        setup_raw = get_file_version(chosen)
        setup_norm = normalize(setup_raw, keep_build=args.keep_build)
        latest_raw = fetch_latest_from_dell()
        latest_norm = normalize(latest_raw, keep_build=args.keep_build)

        # Verdict: only compare setup vs online and print simple result
        if setup_norm and latest_norm:
            cmp = cmp_versions(setup_norm, latest_norm)
            if cmp == 0:
                return Fore.GREEN + (f"✅ | Dell Support Assist | local : v{setup_norm} | online : v{latest_norm}\n")
            elif cmp < 0:
                return Fore.LIGHTRED_EX + (f"❌ | Dell Support Assist | local : v{setup_norm} | online : v{latest_norm}\n")
            else:
                return Fore.YELLOW + (f"⚠️ | Dell Support Assist | version invalide (local: {setup_norm}, online: {latest_norm})\n")
        else:
            return ("ERROR: unable to read setup or online version")
    return main()
def update_version(app_name):
    url = url_commands.get(app_name)
    if not url:
        return "URL non définie"
    if "github" not in url_commands.get(app_name): 
        if app_name == "Firefox":
            try:
                url = urlfirefox
                html = requests.get(url).text
                soup = BeautifulSoup(html, "html.parser")
                version = soup.html["data-latest-firefox"]
                return version
            except ValueError as e:
                return f"Erreur : {e}"
        elif app_name == "Adobe Acrobat Reader":
            try:
                url = urlaar
                html = requests.get(url).text
                soup = BeautifulSoup(html, "html.parser")
                link_next = soup.find("link", rel="next")
                if link_next and "title" in link_next.attrs:
                    title = link_next["title"]
                    match = re.search(r"\d+(\.\d+)+", title)
                    version = match.group(0) if match else "Version introuvable"
                    return version
                else:
                    return "Version introuvable"
            except Exception as e:
                return f"Erreur : {e}"
        elif app_name == "Chrome":
            try:
                url = urlchrome
                feed = feedparser.parse(url)
                latest_post = feed.entries[0]
                content = latest_post.summary
                browser_match = re.search(r"Browser version ([\d\.]+)", content)
                browser_version = browser_match.group(1) if browser_match else None
                browser_version = get_version_without_build(browser_version)
                return browser_version
            except Exception as e:
                return f'Erreur : {e}'
        elif app_name == "AnyDesk":
            try:
                scraper = cloudscraper.create_scraper()
                r = scraper.get(urlanydesk, timeout=20)
                r.raise_for_status()
                soup = BeautifulSoup(r.text, "html.parser")
                m = re.search(r"Version\s+([0-9]+(?:\.[0-9]+)+)", soup.get_text("\n"))
                if m:
                    version = m.group(1)
                    return version
                else:
                    return f'Version Introuvable'
            except Exception as e:
                return f'Erreur : {e}'
        elif app_name == "Malwarebytes":
            try:
                url = urlmalwarebytes

                scraper = cloudscraper.create_scraper()

                articles = []
                next_url = url

                while next_url:
                    r = scraper.get(next_url, timeout=20)
                    r.raise_for_status()
                    data = r.json()

                    articles.extend(data.get("articles", []))
                    next_url = data.get("next_page")

                versions = []

                for article in articles:
                    title = article.get("title", "")
                    link = article.get("html_url", "")

                    if "windows" not in title.lower():
                        continue

                    m = re.search(r"(\d+\.\d+\.\d+\.\d+)", title)
                    if m:
                        versions.append((Version(m.group(1)), link))

                if not versions:
                    return "Aucune version trouvée"
                else:
                    latest = max(versions, key=lambda x: x[0])
                    v = latest[0]
                    return Version(f"{v.major}.{v.minor}.{v.micro}")
            except Exception as e:
                return f'Erreur  {e}'
            
        elif app_name == "ESET Endpoint Security" or app_name == "ESET Full Disk Encryption for Windows":
            try:
                url = urleset
                headers = {
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
                }
                r = requests.get(url, headers=headers, timeout=20)
                r.raise_for_status()
                soup = BeautifulSoup(r.text, "html.parser")

                for heading in soup.find_all(["h2", "h3"]):
                    if heading.get_text(strip=True) == app_name:
                        table = heading.find_next("table")
                        if table:
                            rows = table.find_all("tr")
                            cells = rows[1].find_all("td")
                            version = cells[1].get_text(strip=True)
                            return version.rsplit(".", 1)[0]

                return "Version introuvable"
            except Exception as e:
                return f"Erreur : {e}"
        
        elif app_name == "Microsoft Edge":
            try:
                url = urlmsedge
                headers = {
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
                }
                r = requests.get(url, headers=headers, timeout=20)
                r.raise_for_status()
                soup = BeautifulSoup(r.text, "html.parser")

                for heading in soup.find_all(["h2", "h3"]):
                    text = heading.get_text(strip=True)
                    if text.startswith("Version"):
                        m = re.search(r"(\d+\.\d+\.\d+\.\d+)", text)
                        if m:
                            v = Version(m.group(1))
                            return Version(f"{v.major}.{v.minor}.{v.micro}")
                
                return "Version introuvable"
            except Exception as e:
                return f"Erreur : {e}"
            
        elif app_name == "Veeam Agent For Microsoft Windows":
            try:
                options = Options()
                options.add_argument("--headless=new")
                options.add_argument("--no-sandbox")
                options.add_argument("--disable-dev-shm-usage")
                options.add_argument("--disable-gpu")
                options.add_argument("--disable-extensions")
                options.add_argument("--disable-logging")
                options.add_argument("--log-level=3")
                options.add_argument("--silent")
                options.add_argument("--output=/dev/null")
                options.add_experimental_option("excludeSwitches", ["enable-logging"])
                service = webdriver.chrome.service.Service(log_path=os.devnull)
                driver = webdriver.Chrome(options=options, service=service)
                driver.get("https://www.veeam.com/products/downloads/latest-version.html")

                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                )

                text = driver.find_element(By.TAG_NAME, "body").text
                driver.quit()

                lines = text.split("\n")
                for i, line in enumerate(lines):
                    if "Veeam Agent for Microsoft Windows" in line:
                        for j in range(i, min(i+10, len(lines))):
                            m = re.search(r"(\d+\.\d+\.\d+)", lines[j])
                            if m:
                                return m.group(1)

                return "Version introuvable"
            except Exception as e:
                return f"Error : {e}"

        elif app_name == "Open VPN Connect":
                url = urlopenvpn
                headers = {
                    "User-Agent": (
                        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) "
                        "Chrome/120.0.0.0 Safari/537.36"
                    )
                }

                try:
                    r = requests.get(url, headers=headers, timeout=20)
                    r.raise_for_status()
                except requests.RequestException as e:
                    return f"Erreur réseau : {e}"

                soup = BeautifulSoup(r.text, "html.parser")

                for heading in soup.find_all(["h2", "h3"]):
                    text = heading.get_text(strip=True)
                    m = re.match(r"^(\d+\.\d+\.\d+)\s*\(", text)
                    if m:
                        return Version(m.group(1))

                return "Version introuvable"


    else:
        try:
            resp = requests.get(url, timeout=8)
            resp.raise_for_status()
            data = resp.json()
            tag = data.get("tag_name") or data.get("name")
            if not tag:
                return "Tag non trouvé"
            tag = str(tag).lstrip("vV")
            m = re.search(r"\d+(\.\d+)*", tag)
            return m.group(0) if m else tag
        except requests.RequestException as e:
            return f"Erreur réseau: {e}"
        except ValueError:
            return "Réponse JSON invalide"


def get_version(app_name):
    if app_name == "Malwarebytes":
        return get_malwarebytes_version()
    
    command = install_commands.get(app_name)
    if not command:
        return "Non défini"
    
    try:
        result = subprocess.run(
            ["powershell", "-Command", command],
            capture_output=True, text=True
        )
        version = result.stdout.strip()
        if version:
            return version
        else:
            return "Non installé"
    except Exception as e:
        return f"Erreur: {e}"

for i, option in enumerate(options):
    var = ctk.BooleanVar()
    checkbox = ctk.CTkCheckBox(
        master=frame_options,
        text=option,
        variable=var,
        hover_color="magenta",
        border_color="gray",
        border_width=2,
        corner_radius=6,
        fg_color="darkmagenta",
    )

    row = i // cols
    col = i % cols

    checkbox.grid(row=row, column=col, padx=8, pady=4, sticky="w")
    vars.append(var)

def get_version_without_build(full_version):
    if not full_version:
        return None
    s = str(full_version).strip()
    try:
        v = Version(s)
        parts = [str(v.major)]
        if v.minor is not None:
            parts.append(str(v.minor))
        if v.micro is not None:
            parts.append(str(v.micro))
        return '.'.join(parts)
    except InvalidVersion:
        m = re.search(r"(\d+(?:\.\d+){0,2})", s)
        return m.group(1) if m else None

def check_selection():
    btn.configure(state="disabled", text="Verification en cours...")
    os.system('cls')
    def run():
        for var, option in zip(vars, options):
            if not var.get():
                continue

            if option == "Dell Support Assist":
                dell = get_dell_support_assist()
                print(dell)
                continue

            full_local = get_version(option)      
            local_version_str = get_version_without_build(full_local)

            online_version_str = update_version(app_name=option)

            if not local_version_str:
                if option == "Dell Support Assist":
                    return
                else:
                    print(Fore.YELLOW + f"⚠️ | {option} | local : non installé ou version introuvable ({full_local})\n")
                    continue
            if not online_version_str or "Erreur" in str(online_version_str):
                print(Fore.YELLOW + f"⚠️ | {option} | online : version introuvable ({online_version_str})\n")
                continue

            try:
                local_version = Version(local_version_str)
                online_version = Version(str(online_version_str))
                if online_version > local_version:
                    print(Fore.LIGHTRED_EX + f"❌ | {option} | local : v{local_version} | online : v{online_version}\n")
                else:
                    print(Fore.GREEN + f"✅ | {option} | local : v{local_version} | online : v{online_version}\n")
            except InvalidVersion:
                print(Fore.YELLOW + f"⚠️ | {option} | version invalide (local: {local_version_str}, online: {online_version_str})\n")
        print(Fore.WHITE + "")
        btn.configure(state="normal", text="Verifier les versions")

    threading.Thread(target=run, daemon=True).start()

def toggle_all():
    state = btn_all_var.get()
    for var in vars:
        var.set(state)

btn_all_var = ctk.BooleanVar()
btn_all = ctk.CTkSwitch(app, text="All", variable=btn_all_var, command=toggle_all,
                        button_color="darkmagenta", progress_color="darkmagenta")
btn_all.grid(pady=5)

btn = ctk.CTkButton(app, text="Vérifier les versions", command=check_selection)
btn.grid(pady=10)
app.mainloop()