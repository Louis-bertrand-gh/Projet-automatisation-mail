# -*- coding: utf-8 -*-
"""
Assistant Arrivées Tardives - Camping Huttopia Lac de Rillé
=============================================================

Application de bureau (Windows) permettant à l'accueil du camping d'envoyer
automatiquement, via Outlook, un email d'arrivée tardive à un ou plusieurs
clients, avec en pièce jointe le plan du camping où l'emplacement du client
est indiqué.

Fonctionnement :
    1. L'utilisateur saisit Nom Prénom / N° d'emplacement(s) / Email du client.
    2. Il clique sur "Ajouter à la liste" -> le client est ajouté à une file
       d'attente et le formulaire se vide pour saisir le client suivant.
    3. Une fois tous les clients du soir ajoutés, un clic sur
       "Envoyer tous les emails" déclenche l'envoi automatique via Outlook,
       avec le bon plan (ou la fusion des plans) en pièce jointe.
    4. Chaque envoi (réussi ou en échec) est journalisé dans
       historique_envois.csv à côté de l'application.

Auteur : généré avec Claude (Anthropic) pour le Camping Huttopia Lac de Rillé.
"""

import csv
import datetime
import html
import json
import os
import re
import sys
import subprocess
import tempfile
import threading
import traceback
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ---------------------------------------------------------------------------
# Import optionnel de pywin32 : l'appli doit pouvoir s'ouvrir (pour tester
# l'interface) même sur une machine où pywin32 ne serait pas encore installé.
# L'erreur ne surgira qu'au moment réel de l'envoi.
# ---------------------------------------------------------------------------
try:
    import win32com.client as win32  # type: ignore
    import pythoncom  # type: ignore
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

# ---------------------------------------------------------------------------
# Fusion de PDF (plusieurs emplacements => plusieurs plans à fusionner)
# ---------------------------------------------------------------------------
try:
    from pypdf import PdfWriter, PdfReader  # type: ignore
    PYPDF_AVAILABLE = True
except ImportError:
    PYPDF_AVAILABLE = False


# =============================================================================
# CONFIGURATION / CHEMINS
# =============================================================================

def get_base_dir() -> str:
    """
    Retourne le dossier où se trouve l'exécutable (ou le script en mode dev).
    Essentiel pour un .exe PyInstaller : on ne doit PAS utiliser sys._MEIPASS
    (dossier temporaire d'extraction) pour les données persistantes
    (plans, historique), sinon elles disparaîtraient à chaque lancement.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = get_base_dir()
PLANS_DIR = os.path.join(BASE_DIR, "plans")
HISTORIQUE_PATH = os.path.join(BASE_DIR, "historique_envois.csv")
CAMPING_NOM = "Camping Huttopia Lac de Rillé"
APP_VERSION = "0.4"
APP_VERSION_WIN = "0.4.0.0"
CONFIG_PATH = os.path.join(BASE_DIR, "parametres.json")

os.makedirs(PLANS_DIR, exist_ok=True)


# =============================================================================
# MODELE DE DONNEES
# =============================================================================

@dataclass
class Client:
    nom: str
    emplacements: List[str]
    email: str
    langue: str = "Français"
    langue2: Optional[str] = None
    statut: str = "En attente"
    detail: str = ""
    deja_envoye: bool = False

    def emplacements_str(self) -> str:
        return ", ".join(self.emplacements)


# =============================================================================
# GABARIT DE L'EMAIL
# =============================================================================

def format_subject(client: Client) -> str:
    subjects = {
        "Français": "Votre arrivée tardive",
        "Anglais": "Your Late Arrival",
        "Néerlandais": "Uw late aankomst",
        "Allemand": "Ihre späte Anreise",
        "Espagnol": "Su llegada tardía",
    }
    return f"{subjects.get(client.langue, subjects['Français'])} - {CAMPING_NOM} - {client.nom}"


def _format_single_language_body(client: Client, langue: str) -> str:
    numeros = html.escape(client.emplacements_str())
    camping = html.escape(CAMPING_NOM)
    messages = {
        "Français": """
            <p>Bonsoir,</p>
            <p>Nous sommes très heureux de vous accueillir prochainement au {camping}.</p>
            <p>Pour vous guider au mieux, vous trouverez en pièce jointe le plan du site,
            avec l&rsquo;itinéraire pour rejoindre votre hébergement (numéro {numeros}).</p>
            <p>Afin que votre arrivée tardive se fasse en toute tranquillité :</p>
            <ul><li>Les clés vous attendent directement dans votre hébergement.</li>
            <li>Un classeur avec toutes les informations utiles (infos pratiques,
            fonctionnement du Camping, etc.) est à votre disposition dans l&rsquo;hébergement.</li></ul>
            <p>Pour préserver le calme et l&rsquo;esprit nature du Camping, le site est sans voiture :
            à votre arrivée, merci de vous garer sur le parking. Si besoin, vous pouvez emprunter
            les chariots verts pour transporter vos affaires&nbsp;; nous vous remercions de bien
            vouloir les remettre à leur place une fois votre installation terminée, pour les prochains arrivants.</p>
            <p>Nous vous invitons à passer nous voir le lendemain matin à la réception pour finaliser
            votre arrivée. La réception ouvre à 8h&nbsp;: n&rsquo;hésitez pas à venir quand vous serez
            prêts, nous serons ravis de vous accueillir et de répondre à vos questions.</p>
            <p>En cas d&rsquo;urgence, vous pouvez nous joindre au numéro d&rsquo;astreinte&nbsp;:
            02 47 24 62 97. Et si vous n&rsquo;avez pas de réseau, un interphone extérieur est
            disponible à l&rsquo;entrée de la réception.</p>
            <p>Nous vous souhaitons une très bonne route et une belle arrivée au Camping.</p>
            <p>Belle soirée,</p>
            <p><strong>L&rsquo;équipe Huttopia Rillé</strong><br>{camping}<br>
            <a href="tel:+33247246297">+33 (0)2 47 24 62 97</a><br>1 Lac de Pincemaille<br>
            37340 Rillé<br><a href="https://www.huttopia.com">www.huttopia.com</a></p>
        """,
        "Néerlandais": """
            <p>Goedenavond,</p>
            <p>Wij kijken er erg naar uit om u binnenkort te mogen verwelkomen op {camping}.</p>
            <p>Om u zo goed mogelijk de weg te wijzen, vindt u in de bijlage de plattegrond van het terrein,
            met de route naar uw accommodatie (nummer {numeros}).</p>
            <p>Om uw late aankomst zo aangenaam en zorgeloos mogelijk te laten verlopen:</p>
            <ul><li>De sleutels liggen direct voor u klaar in uw accommodatie.</li>
            <li>Er ligt een map met alle nuttige informatie (praktische info, de werking van de camping, etc.)
            voor u klaar in uw verblijf.</li></ul>
            <p>Om de rust en de sfeer van de natuur te behouden, is de camping autovrij: wij vragen u vriendelijk
            om bij aankomst op de parkeerplaats te parkeren. Indien nodig kunt u de groene karren gebruiken om uw
            spullen te vervoeren; wij verzoeken u vriendelijk deze na het uitpakken weer terug te zetten voor de volgende gasten.</p>
            <p>Wij nodigen u van harte uit om de volgende ochtend even langs te komen bij de receptie om uw aankomst
            af te ronden. De receptie gaat om 8:00 uur open: kom gerust langs wanneer u er klaar voor bent, wij helpen
            u met plezier verder en beantwoorden graag al uw vragen.</p>
            <p>In geval van nood kunt u ons bereiken op het noodnummer: +33 (0)2 47 24 62 97. Mocht u geen bereik
            hebben, dan hangt er een buitenintercom bij de ingang van de receptie.</p>
            <p>Wij wensen u een hele goede reis en een fijne aankomst op de camping.</p>
            <p>Fijne avond,</p>
            <p><strong>Het team van Huttopia Rillé</strong><br>{camping}<br>
            +33 (0)2 47 24 62 97<br>1 Lac de Pincemaille<br>37340 Rillé<br>
            <a href="https://www.huttopia.com">www.huttopia.com</a></p>
        """,
        "Anglais": """
            <p>Good evening,</p>
            <p>We are delighted to welcome you soon to {camping}.</p>
            <p>To help you find your way, please find attached the map of the campsite,
            including the route to your accommodation (number {numeros}).</p>
            <p>To ensure your late arrival goes smoothly and hassle-free:</p>
            <ul><li>The keys will be waiting for you directly inside your accommodation.</li>
            <li>A binder with all useful information (practical details, campsite rules, etc.)
            is available inside your accommodation.</li></ul>
            <p>To preserve the peace and natural spirit of the campsite, the site is car-free: upon arrival,
            please park your vehicle in the main car park. If needed, you are welcome to use the green trolleys
            to transport your luggage; we kindly ask that you return them to their designated spot once you have settled in,
            for the convenience of future arrivals.</p>
            <p>Please feel free to drop by reception the following morning to finalize your check-in. Reception opens
            at 8:00 am: come whenever you are ready, we will be delighted to welcome you and answer any questions you may have.</p>
            <p>In case of an emergency, you can reach us on our emergency number: +33 (0)2 47 24 62 97.
            If you do not have mobile reception, an outdoor intercom is available at the reception entrance.</p>
            <p>We wish you a safe journey and a wonderful arrival at the campsite.</p>
            <p>Have a lovely evening,</p>
            <p><strong>The Huttopia Rillé team</strong><br>{camping}<br>
            +33 (0)2 47 24 62 97<br>1 Lac de Pincemaille<br>37340 Rillé<br>
            <a href="https://www.huttopia.com">www.huttopia.com</a></p>
        """,
        "Allemand": """
            <p>Guten Abend,</p>
            <p>Wir freuen uns sehr, Sie bald auf dem Campingplatz Huttopia Lac de Rillé begrüßen zu dürfen.</p>
            <p>Damit Sie sich gut zurechtfinden, finden Sie im Anhang den Lageplan des Campingplatzes mit dem Weg
            zu Ihrer Unterkunft (Nummer {numeros}).</p>
            <p>Damit Ihre späte Anreise ganz entspannt verläuft:</p>
            <ul><li>Die Schlüssel warten bereits direkt in Ihrer Unterkunft auf Sie.</li>
            <li>Dort liegt auch eine Informationsmappe mit allen nützlichen Hinweisen (praktische Infos,
            Funktionsweise des Campingplatzes usw.) für Sie bereit.</li></ul>
            <p>Um die Ruhe und den naturnahen Charakter des Campingplatzes zu bewahren, ist das Gelände autofrei:
            Bitte parken Sie bei Ihrer Ankunft auf dem Parkplatz. Bei Bedarf können Sie die grünen Transportwagen nutzen,
            um Ihr Gepäck zu transportieren; wir bitten Sie herzlich, diese nach dem Ausladen wieder an ihren Platz zurückzubringen,
            damit sie den nächsten Ankommenden zur Verfügung stehen.</p>
            <p>Wir laden Sie herzlich ein, am nächsten Morgen an der Rezeption vorbeizuschauen, um Ihre Ankunft anzumelden.
            Die Rezeption öffnet um 8:00 Uhr: Kommen Sie einfach vorbei, wenn Sie bereit sind – wir freuen uns darauf,
            Sie willkommen zu heißen und Ihre Fragen zu beantworten.</p>
            <p>Im Notfall erreichen Sie uns unter der Notfallnummer: +33 (0)2 47 24 62 97. Sollten Sie keinen Empfang haben,
            steht Ihnen am Eingang der Rezeption eine Außengegensprechanlage zur Verfügung.</p>
            <p>Wir wünschen Ihnen eine gute Fahrt und ein schönes Ankommen auf dem Campingplatz.</p>
            <p>Einen schönen Abend,</p>
            <p><strong>Das Team von Huttopia Rillé</strong><br>{camping}<br>
            +33 (0)2 47 24 62 97<br>1 Lac de Pincemaille<br>37340 Rillé<br>
            <a href="https://www.huttopia.com">www.huttopia.com</a></p>
        """,
        "Espagnol": """
            <p>Buenas noches:</p>
            <p>Nos complace mucho darles la bienvenida próximamente al {camping}.</p>
            <p>Para guiarles de la mejor manera, encontrarán adjunto el mapa del sitio, con el itinerario para llegar
            a su alojamiento (número {numeros}).</p>
            <p>Para que su llegada tardía transcurra con total tranquilidad:</p>
            <ul><li>Las llaves les estarán esperando directamente en su alojamiento.</li>
            <li>Una carpeta con toda la información útil (datos prácticos, funcionamiento del camping, etc.)
            está a su disposición en el alojamiento.</li></ul>
            <p>Para preservar la tranquilidad y el espíritu natural del camping, el recinto es sin coches: a su llegada,
            les rogamos aparcar en el parking. Si lo necesitan, pueden utilizar los carros verdes para transportar sus pertenencias;
            les agradecemos que los devuelvan a su lugar una vez instalados, para los próximos clientes.</p>
            <p>Les invitamos a pasar a vernos a la mañana siguiente por la recepción para finalizar su llegada.
            La recepción abre a las 8:00 h: no duden en venir cuando estén listos, estaremos encantados de atenderles
            y responder a sus preguntas.</p>
            <p>En caso de emergencia, pueden contactarnos en el número de guardia: +33 (0)2 47 24 62 97.
            Y si no tienen cobertura, hay un interfono exterior disponible a la entrada de la recepción.</p>
            <p>Les deseamos un muy buen viaje y una feliz llegada al camping.</p>
            <p>Que tengan una feliz tarde,</p>
            <p><strong>El equipo de Huttopia Rillé</strong><br>{camping}<br>
            +33 (0)2 47 24 62 97<br>1 Lac de Pincemaille<br>37340 Rillé<br>
            <a href="https://www.huttopia.com">www.huttopia.com</a></p>
        """,
    }
    message = messages.get(langue, messages["Français"])
    return f"""<html><body style="font-family:Calibri,Arial,sans-serif; font-size:11pt;
    color:#1a1a1a; line-height:1.5;">{message.format(camping=camping, numeros=numeros)}</body></html>"""


def format_body_html(client: Client) -> str:
    langues = [client.langue]
    if client.langue2 and client.langue2 != client.langue:
        langues.append(client.langue2)
    sections = [
        _format_single_language_body(client, langue).replace("<html><body", "<section")
        .replace("</body></html>", "</section>")
        for langue in langues
    ]
    return (
        '<html><body style="font-family:Calibri,Arial,sans-serif; font-size:11pt; '
        'color:#1a1a1a; line-height:1.5;">'
        + '<hr style="border:0;border-top:1px solid #d7ddd9;">'.join(sections)
        + "</body></html>"
    )


# =============================================================================
# GESTION DES PLANS (recherche + fusion PDF)
# =============================================================================

def find_plan_for_emplacement(numero: str) -> Optional[str]:
    """Cherche le plan PDF ou image associé à un emplacement."""
    numero = numero.strip()
    extension_priority = {".pdf": 0, ".png": 1, ".jpg": 2, ".jpeg": 3}
    try:
        candidates = []
        for filename in os.listdir(PLANS_DIR):
            stem, extension = os.path.splitext(filename)
            extension = extension.lower()
            if stem.lower() == numero.lower() and extension in extension_priority:
                candidates.append((extension_priority[extension], filename))
        if candidates:
            _, filename = min(candidates)
            return os.path.join(PLANS_DIR, filename)
    except FileNotFoundError:
        pass
    return None


def build_attachment_for_client(client: Client, tmp_dir: str) -> Tuple[Optional[str], List[str]]:
    """
    Retourne (chemin_piece_jointe, liste_emplacements_sans_plan).
    - 1 seul plan trouvé (PDF ou image) -> on attache le fichier original.
    - Plusieurs plans trouvés -> fusion en un PDF temporaire (nécessite pypdf).
    - Aucun plan trouvé -> (None, tous_les_emplacements).
    """
    found = {}
    missing = []
    for emp in client.emplacements:
        path = find_plan_for_emplacement(emp)
        if path:
            found[emp] = path
        else:
            missing.append(emp)

    if not found:
        return None, missing

    if len(found) == 1:
        return next(iter(found.values())), missing

    # Plusieurs PDF peuvent être fusionnés. Les formats image sont utilisables
    # seuls ; pour plusieurs images, on conserve le premier plan et on signale
    # les autres afin de ne jamais envoyer un plan incomplet sans avertissement.
    if not PYPDF_AVAILABLE:
        # Pas de fusion possible : on attache seulement le premier plan trouvé
        # et on prévient l'utilisateur via "missing" (ajout des autres numéros).
        first_path = next(iter(found.values()))
        missing_due_to_merge = [e for e in found if found[e] != first_path]
        return first_path, missing + missing_due_to_merge

    writer = PdfWriter()
    for emp, path in found.items():
        if not path.lower().endswith(".pdf"):
            first_path = next(iter(found.values()))
            missing_due_to_merge = [e for e in found if found[e] != first_path]
            return first_path, missing + missing_due_to_merge
        reader = PdfReader(path)
        for page in reader.pages:
            writer.add_page(page)

    safe_nom = re.sub(r"[^A-Za-z0-9_-]+", "_", client.nom).strip("_") or "client"
    merged_path = os.path.join(tmp_dir, f"plan_{safe_nom}.pdf")
    with open(merged_path, "wb") as f:
        writer.write(f)

    return merged_path, missing


# =============================================================================
# ENVOI VIA OUTLOOK
# =============================================================================

def get_outlook_application():
    """Récupère Outlook déjà ouvert avant de créer une nouvelle instance."""
    try:
        return win32.GetActiveObject("Outlook.Application")
    except Exception:
        pass

    try:
        return win32.Dispatch("Outlook.Application")
    except Exception:
        return win32.DispatchEx("Outlook.Application")


def get_default_outlook_account(namespace, accounts):
    """Retourne le compte lié au magasin Outlook défini par défaut."""
    default_store_id = ""
    try:
        default_store_id = str(namespace.DefaultStore.StoreID)
    except Exception:
        pass

    if default_store_id:
        for index in range(1, accounts.Count + 1):
            try:
                account = accounts.Item(index)
                if str(account.DeliveryStore.StoreID) == default_store_id:
                    return account
            except Exception:
                continue

    if accounts.Count:
        return accounts.Item(1)
    return None


def send_email_via_outlook(client: Client, attachment_path: Optional[str],
                            preview_only: bool = False,
                            outlook_application=None) -> Tuple[bool, str]:
    """
    Envoie (ou affiche en brouillon si preview_only=True) l'email via Outlook.
    Retourne (succes, message_erreur_eventuel).
    """
    if not WIN32_AVAILABLE:
        return False, ("Le module pywin32 n'est pas installé ou Outlook n'est pas "
                        "disponible sur cette machine.")

    com_initialized = False
    try:
        # Chaque thread doit initialiser COM avant tout appel à Outlook.
        if outlook_application is None:
            pythoncom.CoInitialize()
            com_initialized = True
            outlook = get_outlook_application()
        else:
            outlook = outlook_application

        namespace = outlook.GetNamespace("MAPI")
        accounts = namespace.Accounts
        if accounts.Count == 0:
            return False, (
                "Aucun compte Outlook n'est configuré. Ouvrez Outlook classique "
                "et ajoutez le compte du camping."
            )

        # Outlook peut avoir plusieurs comptes, même lorsqu'un compte est
        # défini par défaut dans l'interface. Le préciser sur le message évite
        # que COM tente d'utiliser un ancien profil ou un compte secondaire.
        account = get_default_outlook_account(namespace, accounts)
        if account is None:
            return False, "Impossible d'identifier le compte Outlook par défaut."

        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.SendUsingAccount = account
        mail.To = client.email
        mail.Subject = format_subject(client)
        mail.HTMLBody = format_body_html(client)

        if attachment_path and os.path.isfile(attachment_path):
            mail.Attachments.Add(os.path.abspath(attachment_path))
        elif attachment_path:
            return False, f"Pièce jointe introuvable : {attachment_path}"

        if preview_only:
            mail.Display(False)
        else:
            # Save avant Send rend l'opération plus fiable avec les profils
            # Microsoft 365/Exchange des versions récentes d'Outlook classique.
            mail.Save()
            mail.Send()

        return True, ""
    except Exception as exc:  # pragma: no cover - dépend de l'environnement Windows
        return False, explain_outlook_error(exc)
    finally:
        if com_initialized:
            pythoncom.CoUninitialize()


def explain_outlook_error(exc: Exception) -> str:
    """Transforme l'erreur COM générique d'Outlook en diagnostic exploitable."""
    error_text = str(exc)
    if "Échec de l'opération" not in error_text and "Operation failed" not in error_text:
        return error_text

    return (
        f"{error_text} "
        "Vérifications : utilisez Outlook classique (le nouvel Outlook pour "
        "Windows ne prend pas en charge l'automatisation COM), ouvrez-le une "
        "fois avec le compte du camping, définissez ce compte par défaut, puis "
        "relancez l'application. Si le problème persiste, activez la "
        "prévisualisation Outlook et envoyez le message manuellement."
    )


# =============================================================================
# HISTORIQUE (CSV)
# =============================================================================

def log_envoi(client: Client, succes: bool, erreur: str = "", plan_manquant: str = ""):
    file_exists = os.path.isfile(HISTORIQUE_PATH)
    with open(HISTORIQUE_PATH, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        if not file_exists:
            writer.writerow([
                "Date/Heure", "Nom Prenom", "Emplacements", "Email", "Langue",
                "Statut", "Plans manquants", "Erreur"
            ])
        writer.writerow([
            datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            client.nom,
            client.emplacements_str(),
            client.email,
            client.langue,
            "Envoyé" if succes else "Échec",
            plan_manquant,
            erreur,
        ])


# =============================================================================
# VALIDATION DES SAISIES
# =============================================================================

EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def parse_emplacements(raw: str) -> List[str]:
    """Découpe une saisie libre en liste de numéros d'emplacement.
    Sépare sur virgule, point-virgule, slash ou espace."""
    parts = re.split(r"[,;/\s]+", raw.strip())
    return [p for p in parts if p]


def validate_client_input(nom: str, emplacements_raw: str, email: str) -> Tuple[bool, str]:
    if not nom.strip():
        return False, "Le nom et prénom du client sont obligatoires."
    emplacements = parse_emplacements(emplacements_raw)
    if not emplacements:
        return False, "Au moins un numéro d'emplacement est obligatoire."
    if not email.strip() or not EMAIL_REGEX.match(email.strip()):
        return False, "L'adresse email n'est pas valide."
    return True, ""


# =============================================================================
# INTERFACE GRAPHIQUE (Tkinter)
# =============================================================================

class App(tk.Tk):
    LANGUES = ("Français", "Anglais", "Néerlandais", "Allemand", "Espagnol")
    DRAPEAUX = {
        "Français": "🇫🇷",
        "Anglais": "🇬🇧",
        "Néerlandais": "🇳🇱",
        "Allemand": "🇩🇪",
        "Espagnol": "🇪🇸",
    }

    def __init__(self):
        super().__init__()
        self.title(f"Assistant Arrivées Tardives - Huttopia Lac de Rillé v{APP_VERSION}")
        self.geometry("1120x760")
        self.minsize(760, 620)

        self.clients: List[Client] = []
        self.settings = self._load_settings()
        self._click_count = 0
        self._click_item = None
        self._click_after_id = None

        self._configure_theme()
        self._build_menu()
        self._build_form()
        self._build_liste()
        self._build_footer()

        self._refresh_tree()
        self.after(150, lambda: self.entry_nom.focus_set())

    def _configure_theme(self):
        self.colors = {
            "green": "#006b4f",
            "green_dark": "#004c3a",
            "green_light": "#e7f1ec",
            "beige": "#e9e0d3",
            "gold": "#b48950",
            "ink": "#26332f",
            "muted": "#6b756f",
            "white": "#ffffff",
        }
        self.configure(bg=self.colors["green_light"])
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", font=("Segoe UI", 10), foreground=self.colors["ink"])
        style.configure("Card.TLabelframe", background=self.colors["white"],
                        bordercolor="#d7ddd9", relief="solid")
        style.configure("Card.TLabelframe.Label", background=self.colors["white"],
                        foreground=self.colors["green_dark"], font=("Segoe UI Semibold", 11))
        style.configure("Title.TLabel", background=self.colors["green"],
                        foreground=self.colors["white"], font=("Segoe UI Semibold", 18))
        style.configure("Subtitle.TLabel", background=self.colors["green"],
                        foreground="#d9eee5", font=("Segoe UI", 10))
        style.configure("Accent.TButton", background=self.colors["green"],
                        foreground=self.colors["white"], padding=(16, 9),
                        font=("Segoe UI Semibold", 10))
        style.map("Accent.TButton", background=[("active", self.colors["green_dark"])])
        style.configure("Send.TButton", background=self.colors["green"],
                        foreground=self.colors["white"], padding=(18, 10),
                        font=("Segoe UI Semibold", 10))
        style.map("Send.TButton",
                  background=[("active", self.colors["green_dark"]),
                              ("disabled", "#9aaba4")],
                  foreground=[("disabled", "#edf3f0")])
        style.configure("Treeview", rowheight=34, font=("Segoe UI", 10),
                        background=self.colors["white"], fieldbackground=self.colors["white"])
        style.configure("Treeview.Heading", background=self.colors["green"],
                        foreground=self.colors["white"], font=("Segoe UI Semibold", 10),
                        padding=8)

    # ------------------------------------------------------------------ #
    # MENU
    # ------------------------------------------------------------------ #
    def _build_menu(self):
        menubar = tk.Menu(self)

        menu_fichier = tk.Menu(menubar, tearoff=0)
        menu_fichier.add_command(label="Ouvrir le dossier des plans",
                                  command=self._ouvrir_dossier_plans)
        menu_fichier.add_command(label="Ouvrir l'historique des envois",
                                  command=self._ouvrir_historique)
        menu_fichier.add_separator()
        menu_fichier.add_command(label="Quitter", command=self.destroy)
        menubar.add_cascade(label="Fichier", menu=menu_fichier)

        menu_aide = tk.Menu(menubar, tearoff=0)
        menu_aide.add_command(label="Paramètres", command=self._ouvrir_parametres)
        menu_aide.add_separator()
        menu_aide.add_command(label="À propos", command=self._afficher_a_propos)
        menubar.add_cascade(label="Aide", menu=menu_aide)

        self.config(menu=menubar)

    def _load_settings(self):
        defaults = {"validate_before_send": True, "delete_on_triple_click": False}
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as handle:
                values = json.load(handle)
            defaults.update({key: bool(values[key]) for key in defaults if key in values})
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            pass
        return defaults

    def _save_settings(self):
        with open(CONFIG_PATH, "w", encoding="utf-8") as handle:
            json.dump(self.settings, handle, ensure_ascii=False, indent=2)

    def _ouvrir_parametres(self):
        dialog = tk.Toplevel(self)
        dialog.title("Paramètres")
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.configure(bg=self.colors["green_light"])

        card = ttk.LabelFrame(dialog, text="Préférences d'envoi et de liste",
                              style="Card.TLabelframe", padding=16)
        card.pack(fill="both", expand=True, padx=18, pady=18)
        validate_var = tk.BooleanVar(value=self.settings["validate_before_send"])
        delete_var = tk.BooleanVar(value=self.settings["delete_on_triple_click"])
        ttk.Checkbutton(
            card, text="Valider chaque email dans Outlook avant l'envoi",
            variable=validate_var
        ).pack(anchor="w", pady=6)
        ttk.Checkbutton(
            card, text="Supprimer un client après 3 clics rapides sur sa ligne",
            variable=delete_var
        ).pack(anchor="w", pady=6)
        ttk.Label(
            card, text="Un double-clic ouvre la modification. Le triple-clic "
                      "est désactivé par défaut pour éviter les suppressions accidentelles.",
            wraplength=430, foreground=self.colors["muted"]
        ).pack(anchor="w", pady=(8, 12))

        def save_and_close():
            self.settings["validate_before_send"] = validate_var.get()
            self.settings["delete_on_triple_click"] = delete_var.get()
            self._save_settings()
            self.var_preview.set(validate_var.get())
            dialog.destroy()

        ttk.Button(card, text="Enregistrer", style="Accent.TButton",
                   command=save_and_close).pack(anchor="e")

    def _ouvrir_dossier_plans(self):
        os.makedirs(PLANS_DIR, exist_ok=True)
        os.startfile(PLANS_DIR) if hasattr(os, "startfile") else None

    def _ouvrir_historique(self):
        if os.path.isfile(HISTORIQUE_PATH):
            if hasattr(os, "startfile"):
                os.startfile(HISTORIQUE_PATH)
        else:
            messagebox.showinfo("Historique", "Aucun envoi n'a encore été enregistré.")

    def _afficher_a_propos(self):
        messagebox.showinfo(
            "À propos",
            "Assistant Arrivées Tardives\n"
            f"{CAMPING_NOM}\n"
            f"Version {APP_VERSION}\n\n"
            "Envoie automatiquement, via Outlook, l'email d'arrivée tardive "
            "et le plan du camping aux clients concernés."
        )

    # ------------------------------------------------------------------ #
    # FORMULAIRE DE SAISIE
    # ------------------------------------------------------------------ #
    def _build_form(self):
        banner = tk.Frame(self, bg=self.colors["green"], height=88)
        banner.pack(fill="x")
        banner.pack_propagate(False)
        ttk.Button(
            banner,
            text="Fermer Outlook",
            command=self._fermer_outlook,
            style="Send.TButton",
        ).pack(side="right", padx=20, pady=24)
        ttk.Label(banner, text="ARRIVÉES TARDIVES", style="Title.TLabel").pack(
            anchor="w", padx=26, pady=(15, 0))
        ttk.Label(banner, text=f"{CAMPING_NOM}  •  Version {APP_VERSION}",
                  style="Subtitle.TLabel").pack(anchor="w", padx=28)

        frame = ttk.LabelFrame(self, text="Ajouter un client à la liste",
                               style="Card.TLabelframe", padding=12)
        frame.pack(fill="x", padx=20, pady=(18, 8))

        ttk.Label(frame, text="Nom Prénom :").grid(row=0, column=0, sticky="e", padx=(4, 3), pady=6)
        self.var_nom = tk.StringVar()
        self.entry_nom = ttk.Entry(frame, textvariable=self.var_nom, width=35)
        self.entry_nom.grid(
            row=0, column=1, sticky="ew", padx=(0, 2), pady=6)
        ttk.Button(frame, text="×", width=2,
                   command=lambda: self.var_nom.set("")).grid(
                       row=0, column=2, padx=(0, 8), pady=6)

        ttk.Label(frame, text="N° emplacement(s) :").grid(
            row=0, column=3, sticky="e", padx=(4, 3), pady=6)
        self.var_emplacements = tk.StringVar()
        self.entry_emplacements = ttk.Entry(frame, textvariable=self.var_emplacements, width=20)
        self.entry_emplacements.grid(
            row=0, column=4, sticky="ew", padx=(0, 2), pady=6)
        ttk.Button(frame, text="×", width=2,
                   command=lambda: self.var_emplacements.set("")).grid(
                       row=0, column=5, padx=(0, 8), pady=6)
        ttk.Label(frame, text="(ex : 12 ou 12, 13)", foreground="#666").grid(
            row=1, column=4, columnspan=2, sticky="w", padx=(0, 10), pady=(0, 4))

        ttk.Label(frame, text="Email du client :").grid(
            row=0, column=6, sticky="e", padx=(4, 3), pady=6)
        self.var_email = tk.StringVar()
        self.entry_email = ttk.Entry(frame, textvariable=self.var_email, width=30)
        self.entry_email.grid(
            row=0, column=7, sticky="ew", padx=(0, 2), pady=6)
        ttk.Button(frame, text="×", width=2,
                   command=lambda: self.var_email.set("")).grid(
                       row=0, column=8, padx=(0, 8), pady=6)

        ttk.Label(frame, text="Langue :").grid(
            row=0, column=9, sticky="e", padx=(4, 3), pady=6)
        self.var_langue = tk.StringVar(value=self.LANGUES[0])
        self.combo_langue = ttk.Combobox(
            frame, textvariable=self.var_langue,
            values=[f"{self.DRAPEAUX[langue]}  {langue}" for langue in self.LANGUES],
            state="readonly", width=17
        )
        self.combo_langue.current(0)
        self.combo_langue.grid(row=0, column=10, sticky="ew", padx=(0, 2), pady=6)
        self.btn_ajouter_langue = ttk.Button(
            frame, text="+", width=3, command=self._show_second_language
        )
        self.btn_ajouter_langue.grid(row=1, column=10, sticky="w", padx=(0, 2), pady=2)

        self.label_langue2 = ttk.Label(frame, text="2e langue :")
        self.label_langue2.grid(row=2, column=9, sticky="e", padx=(4, 3), pady=4)
        self.var_langue2 = tk.StringVar(value="Aucune")
        self.combo_langue2 = ttk.Combobox(
            frame, textvariable=self.var_langue2,
            values=["Aucune"] + [
                f"{self.DRAPEAUX[langue]}  {langue}" for langue in self.LANGUES
            ],
            state="readonly", width=17
        )
        self.combo_langue2.current(0)
        self.combo_langue2.grid(row=2, column=10, sticky="ew", padx=(0, 2), pady=4)
        self.btn_retirer_langue = ttk.Button(
            frame, text="−", width=3, command=self._hide_second_language
        )
        self.btn_retirer_langue.grid(row=2, column=11, padx=(0, 4), pady=4)
        self.label_langue2.grid_remove()
        self.combo_langue2.grid_remove()
        self.btn_retirer_langue.grid_remove()

        self.btn_frame = ttk.Frame(frame)
        self.btn_frame.grid(row=3, column=0, columnspan=12, sticky="ew", padx=4, pady=(6, 8))

        ttk.Button(self.btn_frame, text="➕ Ajouter à la liste",
                   command=self._ajouter_client, style="Accent.TButton").pack(
                       side="left", padx=(0, 8))
        ttk.Button(self.btn_frame, text="Vider le formulaire",
                   command=self._vider_formulaire).pack(side="left")

        for i in (1, 4, 7, 10):
            frame.grid_columnconfigure(i, weight=1, uniform="field")
        for i in (0, 2, 3, 5, 6, 8, 9):
            frame.grid_columnconfigure(i, weight=0)

        for entry in (self.entry_nom, self.entry_emplacements, self.entry_email):
            self._bind_word_delete(entry)

        # Ajout rapide avec Entrée
        self.bind("<Return>", lambda e: self._ajouter_client())

    def _fermer_outlook(self):
        if not messagebox.askyesno(
            "Fermer Outlook",
            "Fermer toutes les instances d'Outlook ?\n"
            "Les messages ou modifications non enregistrés seront perdus.",
            parent=self,
        ):
            return

        try:
            result = subprocess.run(
                ["taskkill", "/F", "/IM", "OUTLOOK.EXE", "/T"],
                capture_output=True,
                text=True,
                check=False,
            )
        except OSError as exc:
            messagebox.showerror(
                "Fermeture impossible",
                f"Windows n'a pas pu fermer Outlook : {exc}",
                parent=self,
            )
            return

        if result.returncode == 0:
            self._set_status("Toutes les instances d'Outlook ont été fermées.")
            messagebox.showinfo(
                "Outlook fermé",
                "Toutes les instances d'Outlook ont été fermées.",
                parent=self,
            )
        elif result.returncode == 128:
            self._set_status("Aucune instance d'Outlook n'est actuellement ouverte.")
            messagebox.showinfo(
                "Outlook",
                "Aucune instance d'Outlook n'est actuellement ouverte.",
                parent=self,
            )
        else:
            details = result.stderr.strip() or result.stdout.strip()
            messagebox.showerror(
                "Fermeture impossible",
                "Windows n'a pas pu fermer toutes les instances d'Outlook."
                + (f"\n\nDétail : {details}" if details else ""),
                parent=self,
            )

    def _bind_word_delete(self, entry):
        entry.bind("<Control-BackSpace>", lambda event: self._delete_word(entry, backward=True))
        entry.bind("<Control-Delete>", lambda event: self._delete_word(entry, backward=False))

    @staticmethod
    def _delete_word(entry, backward):
        if entry.selection_present():
            start = entry.index(tk.SEL_FIRST)
            end = entry.index(tk.SEL_LAST)
            entry.delete(start, end)
            return "break"
        cursor = entry.index(tk.INSERT)
        text = entry.get()
        if backward:
            match = re.search(r"\S+\s*$", text[:cursor])
            if match:
                entry.delete(match.start(), cursor)
        else:
            match = re.search(r"\s*\S+", text[cursor:])
            if match:
                entry.delete(cursor, cursor + match.end())
        return "break"

    def _clear_language(self):
        self.var_langue.set(self.LANGUES[0])
        self.combo_langue.current(0)

    def _show_second_language(self):
        self.btn_ajouter_langue.grid_remove()
        self.label_langue2.grid()
        self.combo_langue2.grid()
        self.btn_retirer_langue.grid()

    def _hide_second_language(self):
        self._clear_second_language()
        self.label_langue2.grid_remove()
        self.combo_langue2.grid_remove()
        self.btn_retirer_langue.grid_remove()
        self.btn_ajouter_langue.grid()

    def _clear_second_language(self):
        self.var_langue2.set("Aucune")
        self.combo_langue2.current(0)

    def _selected_language(self, value):
        return value.split("  ", 1)[-1] if value else self.LANGUES[0]

    def _vider_formulaire(self):
        self.var_nom.set("")
        self.var_emplacements.set("")
        self.var_email.set("")
        self._clear_language()
        self._clear_second_language()
        self._hide_second_language()

    def _ajouter_client(self):
        nom = self.var_nom.get()
        emplacements_raw = self.var_emplacements.get()
        email = self.var_email.get()
        langue = self._selected_language(self.var_langue.get())
        langue2_value = self.var_langue2.get()
        langue2 = None if langue2_value == "Aucune" else self._selected_language(langue2_value)

        ok, message = validate_client_input(nom, emplacements_raw, email)
        if not ok:
            messagebox.showwarning("Saisie incomplète", message)
            return

        client = Client(
            nom=nom.strip(),
            emplacements=parse_emplacements(emplacements_raw),
            email=email.strip(),
            langue=langue,
            langue2=langue2,
        )

        # Vérification immédiate de la disponibilité du/des plan(s)
        plan_paths = {
            emplacement: find_plan_for_emplacement(emplacement)
            for emplacement in client.emplacements
        }
        found_any = any(plan_paths.values())
        missing = [e for e, path in plan_paths.items() if path is None]
        if not found_any:
            client.statut = "⚠ Aucun plan trouvé"
        elif missing:
            client.statut = f"⚠ Plan manquant : {', '.join(missing)}"
        else:
            client.statut = "Prêt à envoyer"

        self.clients.append(client)
        self._refresh_tree()
        self._vider_formulaire()
        self._set_status(f"{len(self.clients)} client(s) dans la liste.")

    # ------------------------------------------------------------------ #
    # LISTE DES CLIENTS AJOUTÉS
    # ------------------------------------------------------------------ #
    def _build_liste(self):
        frame = ttk.LabelFrame(self, text="Clients à traiter  •  Double-clic : modifier",
                               style="Card.TLabelframe", padding=8)
        frame.pack(fill="both", expand=True, padx=20, pady=8)

        columns = ("nom", "emplacements", "email", "langue", "statut", "supprimer")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", height=10)
        self.tree.tag_configure("deja_envoye", foreground="#666666")
        self.tree.heading("nom", text="Nom Prénom")
        self.tree.heading("emplacements", text="Emplacement(s)")
        self.tree.heading("email", text="Email")
        self.tree.heading("langue", text="Langue")
        self.tree.heading("statut", text="Statut")
        self.tree.heading("supprimer", text="")

        self.tree.column("nom", width=190, minwidth=110, stretch=True)
        self.tree.column("emplacements", width=135, minwidth=90, stretch=True)
        self.tree.column("email", width=250, minwidth=150, stretch=True)
        self.tree.column("langue", width=190, minwidth=120, stretch=True)
        self.tree.column("statut", width=220, minwidth=130, stretch=True)
        self.tree.column("supprimer", width=38, minwidth=38, stretch=False, anchor="center")

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=8)
        scrollbar.grid(row=0, column=1, sticky="ns", pady=8)
        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<Delete>", lambda event: self._supprimer_selection())
        self.tree.bind("<Control-Delete>", lambda event: self._supprimer_selection())

        action_frame = ttk.Frame(frame)
        action_frame.grid(row=0, column=2, sticky="ns", padx=(10, 8), pady=8)

        ttk.Button(action_frame, text="Supprimer la sélection",
                   command=self._supprimer_selection).pack(fill="x", pady=(0, 6))
        ttk.Button(action_frame, text="Autoriser le renvoi",
                   command=self._autoriser_renvoi_selection).pack(fill="x", pady=(0, 6))
        ttk.Button(action_frame, text="Vider toute la liste",
                   command=self._vider_liste).pack(fill="x")

    def _on_tree_click(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        self.tree.selection_set(item)
        if self.tree.identify_column(event.x) == "#6":
            self._supprimer_client(int(item))
            return
        if self._click_after_id is not None:
            self.after_cancel(self._click_after_id)
        if item != self._click_item:
            self._click_count = 0
        self._click_item = item
        self._click_count += 1
        self._click_after_id = self.after(420, self._process_tree_clicks)

    def _process_tree_clicks(self):
        item = self._click_item
        count = self._click_count
        self._click_count = 0
        self._click_item = None
        self._click_after_id = None
        if not item or not self.tree.exists(item):
            return
        index = int(item)
        if count == 2:
            self._modifier_client(index)
        elif count >= 3 and self.settings["delete_on_triple_click"]:
            if messagebox.askyesno(
                "Supprimer le client",
                f"Supprimer {self.clients[index].nom} de la liste ?"
            ):
                del self.clients[index]
                self._refresh_tree()
                self._set_status(f"{len(self.clients)} client(s) dans la liste.")

    def _modifier_client(self, index: int):
        client = self.clients[index]
        dialog = tk.Toplevel(self)
        dialog.title("Modifier le client")
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.configure(bg=self.colors["green_light"])
        card = ttk.LabelFrame(dialog, text="Informations client",
                              style="Card.TLabelframe", padding=16)
        card.pack(fill="both", expand=True, padx=18, pady=18)

        values = [tk.StringVar(value=client.nom),
                  tk.StringVar(value=client.emplacements_str()),
                  tk.StringVar(value=client.email),
                  tk.StringVar(value=client.langue),
                  tk.StringVar(value=client.langue2 or "Aucune")]
        labels = ("Nom Prénom :", "N° emplacement(s) :", "Email :",
                  "Langue :", "2e langue :")
        for row, (label, value) in enumerate(zip(labels, values)):
            ttk.Label(card, text=label).grid(row=row, column=0, sticky="w", padx=6, pady=7)
            if row in (3, 4):
                field = ttk.Combobox(
                    card, textvariable=value,
                    values=(["Aucune"] if row == 4 else []) + [
                        f"{self.DRAPEAUX[langue]}  {langue}" for langue in self.LANGUES
                    ],
                    state="readonly", width=39
                )
                if row == 3:
                    field.current(self.LANGUES.index(client.langue)
                                  if client.langue in self.LANGUES else 0)
                else:
                    field.current(self.LANGUES.index(client.langue2) + 1
                                  if client.langue2 in self.LANGUES else 0)
            else:
                field = ttk.Entry(card, textvariable=value, width=42)
                self._bind_word_delete(field)
            field.grid(row=row, column=1, padx=6, pady=7)

        def save_client():
            ok, message = validate_client_input(values[0].get(), values[1].get(), values[2].get())
            if not ok:
                messagebox.showwarning("Saisie invalide", message, parent=dialog)
                return
            client.nom = values[0].get().strip()
            client.emplacements = parse_emplacements(values[1].get())
            client.email = values[2].get().strip()
            client.langue = self._selected_language(values[3].get())
            client.langue2 = (None if values[4].get() == "Aucune"
                              else self._selected_language(values[4].get()))
            missing = [e for e in client.emplacements if not find_plan_for_emplacement(e)]
            client.statut = ("⚠ Aucun plan trouvé" if len(missing) == len(client.emplacements)
                             else f"⚠ Plan manquant : {', '.join(missing)}" if missing
                             else "Prêt à envoyer")
            self._refresh_tree()
            dialog.destroy()

        ttk.Button(card, text="Enregistrer", style="Accent.TButton",
                   command=save_client).grid(row=5, column=1, sticky="e", padx=6, pady=(12, 0))

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for idx, client in enumerate(self.clients):
            langues = self.DRAPEAUX.get(client.langue, "") + " " + client.langue
            if client.langue2:
                langues += " / " + self.DRAPEAUX.get(client.langue2, "") + " " + client.langue2
            tags = ("deja_envoye",) if client.deja_envoye else ()
            self.tree.insert("", "end", iid=str(idx), values=(
                client.nom, client.emplacements_str(), client.email,
                langues,
                client.statut, "×"
            ), tags=tags)

    def _supprimer_client(self, index):
        if 0 <= index < len(self.clients):
            del self.clients[index]
            self._refresh_tree()
            self._set_status(f"{len(self.clients)} client(s) dans la liste.")

    def _supprimer_selection(self):
        selected = self.tree.selection()
        if not selected:
            return
        indices = sorted((int(i) for i in selected), reverse=True)
        for i in indices:
            del self.clients[i]
        self._refresh_tree()
        self._set_status(f"{len(self.clients)} client(s) dans la liste.")

    def _autoriser_renvoi_selection(self):
        selected = self.tree.selection()
        if not selected:
            return
        for item in selected:
            index = int(item)
            self.clients[index].deja_envoye = False
            if self.clients[index].statut.startswith("✓"):
                self.clients[index].statut = "Prêt à envoyer"
        self._refresh_tree()
        self._set_status("Le renvoi est autorisé pour la sélection.")

    def _vider_liste(self):
        if not self.clients:
            return
        if messagebox.askyesno("Confirmation", "Vider toute la liste des clients ?"):
            self.clients.clear()
            self._refresh_tree()
            self._set_status("Liste vidée.")

    # ------------------------------------------------------------------ #
    # PIED DE PAGE : options + envoi + journal
    # ------------------------------------------------------------------ #
    def _build_footer(self):
        self.status_bar = ttk.Label(self, text="Prêt.", anchor="w",
                                    background=self.colors["green_dark"],
                                    foreground=self.colors["white"], padding=7)
        self.status_bar.pack(side="bottom", fill="x")

        frame = tk.Frame(self, bg=self.colors["white"], height=54,
                         highlightbackground="#cbd8d2", highlightthickness=1)
        frame.pack(side="bottom", fill="x", padx=20, pady=(0, 8))
        frame.pack_propagate(False)

        self.var_preview = tk.BooleanVar(value=self.settings["validate_before_send"])
        ttk.Checkbutton(
            frame,
            text="Valider chaque email dans Outlook avant envoi",
            variable=self.var_preview,
            command=self._preview_changed,
        ).pack(side="left")

        self.btn_envoyer = ttk.Button(
            self.btn_frame, text="Envoyer tous les emails", command=self._lancer_envoi_tous,
            style="Send.TButton", width=25
        )
        self.btn_envoyer.pack(side="right", padx=(10, 0), pady=2)

        log_frame = ttk.LabelFrame(self, text="Journal des envois",
                                   style="Card.TLabelframe", padding=6)
        log_frame.pack(fill="both", expand=False, padx=20, pady=(0, 8))
        self.txt_log = tk.Text(log_frame, height=8, state="disabled", wrap="word")
        self.txt_log.pack(fill="both", expand=True, padx=6, pady=6)

        if not WIN32_AVAILABLE:
            self._log("⚠ pywin32 n'est pas installé : l'envoi réel via Outlook sera "
                       "impossible tant que la dépendance n'est pas installée "
                       "(voir requirements.txt).")
        if not PYPDF_AVAILABLE:
            self._log("ℹ pypdf n'est pas installé : la fusion automatique de plusieurs "
                       "plans pour un même client ne sera pas disponible.")

    def _preview_changed(self):
        self.settings["validate_before_send"] = self.var_preview.get()
        self._save_settings()

    def _set_status(self, text: str):
        if threading.current_thread() is not threading.main_thread():
            self.after(0, self._set_status, text)
            return
        self.status_bar.config(text=text)

    def _log(self, text: str):
        if threading.current_thread() is not threading.main_thread():
            self.after(0, self._log, text)
            return
        self.txt_log.config(state="normal")
        horodatage = datetime.datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert("end", f"[{horodatage}] {text}\n")
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")

    # ------------------------------------------------------------------ #
    # ENVOI DE TOUS LES EMAILS
    # ------------------------------------------------------------------ #
    def _lancer_envoi_tous(self):
        clients_a_envoyer = [client for client in self.clients if not client.deja_envoye]
        if not clients_a_envoyer:
            messagebox.showinfo(
                "Aucun nouvel envoi",
                "Tous les clients de la liste ont déjà reçu leur email."
            )
            return

        nb = len(clients_a_envoyer)
        if not messagebox.askyesno(
            "Confirmation d'envoi",
            f"Confirmez-vous l'envoi de {nb} nouveau(x) email(s) via Outlook ?"
        ):
            return

        self.btn_envoyer.config(state="disabled")
        self._set_status("Envoi en cours...")
        preview_only = self.var_preview.get()

        # Exécution dans un thread séparé pour ne pas geler l'interface,
        # tout en gardant les appels Outlook (COM) cohérents.
        thread = threading.Thread(
            target=self._envoyer_tous_worker,
            args=(preview_only,),
            daemon=True,
        )
        thread.start()

    def _envoyer_tous_worker(self, preview_only: bool):
        succes_count = 0
        echec_count = 0

        # Une seule connexion COM est utilisée pour toute la session d'envoi.
        # Cela évite de relancer le choix de profil et permet de réutiliser
        # Outlook déjà ouvert en arrière-plan.
        outlook = None
        outlook_error = ""
        com_initialized = False
        if WIN32_AVAILABLE:
            try:
                pythoncom.CoInitialize()
                com_initialized = True
                outlook = get_outlook_application()
            except Exception as exc:
                outlook_error = explain_outlook_error(exc)

        try:
            with tempfile.TemporaryDirectory() as tmp_dir:
                for idx, client in enumerate(list(self.clients)):
                    if client.deja_envoye:
                        self._log(f"  • Email déjà envoyé pour {client.nom}, envoi ignoré.")
                        continue

                    if outlook is None:
                        self._log(f"  ✗ Connexion Outlook impossible : {outlook_error}")
                        client.statut = "✗ Échec Outlook"
                        log_envoi(client, False, outlook_error)
                        echec_count += 1
                        self._update_tree_row(idx, client)
                        continue

                    self._log(f"Traitement de {client.nom} "
                              f"(emplacement {client.emplacements_str()})...")

                    attachment_path, missing = build_attachment_for_client(client, tmp_dir)

                    if missing:
                        self._log(f"  ⚠ Plan(s) introuvable(s) pour : {', '.join(missing)}")

                    if attachment_path is None:
                        self._log(f"  ✗ Aucun plan disponible : email NON envoyé pour "
                                  f"{client.nom}.")
                        client.statut = "✗ Échec (aucun plan)"
                        log_envoi(client, False, "Aucun plan trouvé", ", ".join(missing))
                        echec_count += 1
                        self._update_tree_row(idx, client)
                        continue

                    succes, erreur = send_email_via_outlook(
                        client, attachment_path, preview_only, outlook
                    )

                    if succes:
                        action = "affiché dans Outlook (à valider manuellement)" if preview_only else "envoyé"
                        self._log(f"  ✓ Email {action} pour {client.nom} <{client.email}>.")
                        client.statut = "✓ Envoyé" if not preview_only else "✓ Ouvert dans Outlook"
                        if not preview_only:
                            client.deja_envoye = True
                        succes_count += 1
                    else:
                        self._log(f"  ✗ Échec pour {client.nom} : {erreur}")
                        client.statut = "✗ Échec"
                        echec_count += 1

                    log_envoi(client, succes, erreur, ", ".join(missing))
                    self._update_tree_row(idx, client)
        finally:
            if com_initialized:
                pythoncom.CoUninitialize()

        def finish_on_ui_thread():
            self._log(f"Terminé : {succes_count} succès, {echec_count} échec(s).")
            self._set_status(f"Terminé : {succes_count} succès, {echec_count} échec(s).")
            self.btn_envoyer.config(state="normal")

            if echec_count:
                messagebox.showwarning(
                    "Envoi terminé avec erreurs",
                    f"{succes_count} email(s) envoyé(s), {echec_count} échec(s).\n"
                    "Consultez le journal et l'historique pour le détail."
                )
            else:
                messagebox.showinfo(
                    "Envoi terminé",
                    f"Les {succes_count} email(s) ont bien été traités."
                )

        self.after(0, finish_on_ui_thread)

    def _update_tree_row(self, idx: int, client: Client):
        # Les appels UI depuis un thread secondaire doivent passer par after()
        langues = self.DRAPEAUX.get(client.langue, "") + " " + client.langue
        if client.langue2:
            langues += " / " + self.DRAPEAUX.get(client.langue2, "") + " " + client.langue2
        self.after(0, lambda: self.tree.item(
            str(idx),
            values=(
                client.nom, client.emplacements_str(), client.email,
                langues, client.statut, "×"
            ),
            tags=("deja_envoye",) if client.deja_envoye else ()
        ))


# =============================================================================
# POINT D'ENTREE
# =============================================================================

def main():
    try:
        app = App()
        app.mainloop()
    except Exception:
        # Filet de sécurité : si l'appli plante, on affiche l'erreur au lieu
        # de fermer la fenêtre noire sans explication (utile en .exe).
        traceback.print_exc()
        try:
            messagebox.showerror("Erreur fatale", traceback.format_exc())
        except Exception:
            pass


if __name__ == "__main__":
    main()
