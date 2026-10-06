# Assistant Arrivées Tardives — Camping Huttopia Lac de Rillé

**Version actuelle : 0.4**

Application Windows (.exe) pour envoyer automatiquement, via Outlook, l'email
d'arrivée tardive avec le plan du camping en pièce jointe.

---

## 1. Installation (à faire une seule fois)

**Pré-requis sur le PC de l'accueil :**
- Windows avec **Microsoft Outlook installé et configuré** avec l'adresse mail
  du camping (déjà en place chez vous).
- **Python 3.10+** installé (gratuit) : https://www.python.org/downloads/
  → Important : cocher la case **"Add Python to PATH"** pendant l'installation.

**Étapes :**
1. Copiez tout le dossier `camping-app` sur le PC de l'accueil (clé USB, mail, etc.).
2. Double-cliquez sur `build.bat`.
   - Le script installe automatiquement tout ce qu'il faut, puis fabrique le
     fichier `.exe`.
   - Cela prend 1 à 3 minutes, uniquement la première fois (ou si le code est
     modifié plus tard).
3. Une fois terminé, allez dans le dossier `dist` : vous y trouverez
   `AssistantArriveesTardives.exe`.
4. Copiez ce `.exe`, ainsi que le dossier `plans` (voir ci-dessous), où vous
   voulez sur le PC (par exemple sur le Bureau), **dans le même dossier**.

Vous n'aurez ensuite **plus jamais besoin de relancer `build.bat`**, sauf si
je vous fournis une nouvelle version du code.

## Version et procédure après une modification

Chaque correction ou évolution fonctionnelle augmente la version de **0.1** :
`0.1` devient `0.2`, puis `0.3`, etc.

Pour publier une modification :

1. Modifier `APP_VERSION` et `APP_VERSION_WIN` dans `main.py`.
2. Modifier les trois valeurs de version (`filevers`, `prodvers`,
   `FileVersion` et `ProductVersion`) dans `version_info.txt`.
   La version Windows utilise quatre nombres : par exemple `0.2` devient
   `0.2.0.0`.
3. Exécuter `build.bat`.
4. Tester d'abord le code Python, puis l'`.exe` nouvellement créé dans
   `dist\AssistantArriveesTardives.exe`.
5. Copier ensemble l'`.exe` et le dossier `plans` vers le PC de l'accueil.

### Vérifier que l'`.exe` est la bonne version

- Ouvrir l'application et choisir **Aide → À propos** : la version doit être
  affichée, par exemple `Version 0.2`.
- Vérifier aussi le titre de la fenêtre, qui doit terminer par `v0.2`.
- Dans Windows, clic droit sur l'`.exe` → **Propriétés** → **Détails** :
  `Version du fichier` et `Version du produit` doivent correspondre à
  `0.2.0.0`.
- Toujours vérifier la date de modification du fichier dans `dist` avant de
  le copier sur le PC de l'accueil.

### Que faire après un changement de code pour tester

1. Fermer l'application actuellement ouverte.
2. Exécuter `python -m py_compile main.py` pour détecter les erreurs de syntaxe.
3. Exécuter `build.bat` pour reconstruire un nouvel `.exe`.
4. Ouvrir le nouvel `.exe` depuis `dist`, contrôler la version dans
   **Aide → À propos**, puis tester avec la case de prévisualisation Outlook
   cochée.
5. Vérifier le nom, l'adresse, l'objet et le plan dans Outlook sans envoyer
   immédiatement. Pour un test réel, utiliser une adresse interne de test.
6. Contrôler `historique_envois.csv`, puis effectuer un seul envoi réel validé.

Ne testez pas directement l'envoi automatique sur une liste de clients réelle :
commencez par un client de test, un emplacement dont le plan est connu et une
adresse email contrôlée.

## 1 bis. Compatibilité Outlook et erreur « Échec de l'opération »

L'envoi avec pièce jointe utilise l'automatisation COM officielle d'**Outlook
classique**. Elle reste compatible avec les versions récentes d'Outlook
classique incluses dans Microsoft 365 : l'application initialise COM dans son
thread d'envoi, ouvre un profil MAPI, vérifie qu'un compte est configuré,
enregistre le message puis l'envoie.

Le **nouvel Outlook pour Windows** est une application différente et ne prend
pas en charge cette automatisation COM. Si cette version est utilisée, l'erreur
peut apparaître sous la forme :
`(-2147352567 ... Échec de l'opération.)`

Dans ce cas :

1. Ouvrir Outlook.
2. Désactiver **Utiliser le nouvel Outlook** pour revenir à **Outlook
   classique** (ou installer Outlook classique depuis Microsoft 365).
3. Ouvrir Outlook classique une fois et vérifier que le compte du camping est
   présent et défini par défaut.
4. Fermer puis relancer l'application.

Une mise à jour d'Outlook classique ne nécessite normalement pas de changement
du programme. Après une mise à jour majeure, refaire un test avec la
prévisualisation activée. Si l'entreprise impose le nouvel Outlook, une future
version devra utiliser Microsoft Graph avec une application Microsoft Entra ID
et une autorisation administrateur : ce n'est pas interchangeable avec COM et
ne peut pas être activé de manière fiable sans ces identifiants.

---

## 2. Interface et paramètres

L'application s'ouvre directement avec le curseur dans le champ **Nom
Prénom**. La liste des clients fonctionne ainsi :

- un double-clic sur une ligne ouvre la fenêtre de modification ;
- un triple-clic peut supprimer la ligne uniquement si l'option est activée
  dans **Aide → Paramètres** ;
- la validation avant envoi est activée par défaut et peut être changée dans
  **Aide → Paramètres** ou avec la case située au-dessus du bouton d'envoi.

Les paramètres sont conservés dans `parametres.json`, à côté de l'application.
L'interface utilise une présentation en cartes, une palette verte/beige et la
typographie Segoe UI pour faciliter la lecture à l'accueil. Les champs du
formulaire s'étendent avec la fenêtre, les colonnes du tableau se redimensionnent
et la liste occupe automatiquement tout l'espace disponible. Le libellé de
chaque saisie reste directement collé à son champ.

## 3. Ajouter les plans du camping

Dans le dossier `plans/`, placez un fichier PDF ou une image (PNG/JPG)
**par numéro d'emplacement**, nommé exactement comme le numéro :

```
plans/
 ├── 12.pdf
 ├── 13.pdf
 ├── 45.pdf
 └── ...
```

Chaque fichier doit déjà contenir le plan du camping avec l'emplacement entouré
et le chemin d'accès tracé (comme vous me les préparerez). Si plusieurs formats
existent pour un même numéro, le PDF est prioritaire.

- Si un client réserve **un seul emplacement** → l'appli attache directement
  le PDF ou l'image correspondant.
- Si un client a **plusieurs emplacements** → l'appli fusionne automatiquement
  les plans correspondants en un seul PDF joint (aucune action nécessaire de
  votre part).
- Si un plan est manquant pour un numéro, l'appli vous **prévient avant
  l'envoi** dans la liste et dans le journal (colonne "Statut").

*(Fonctionnalité prévue plus tard : possibilité de joindre manuellement un
plan personnalisé fait à la main pour un client avec plusieurs hébergements —
à coder dans une prochaine version.)*

---

## 4. Utilisation quotidienne

1. Double-cliquez sur `AssistantArriveesTardives.exe`.
2. Pour chaque client en arrivée tardive :
   - Renseignez **Nom Prénom**, **N° d'emplacement(s)** (ex : `12` ou `12, 13`)
     et **Email**.
   - Cliquez sur **"➕ Ajouter à la liste"** (ou appuyez sur Entrée).
   - Le formulaire se vide : vous pouvez saisir le client suivant immédiatement.
3. La liste du bas affiche tous les clients ajoutés, avec un statut
   (« Prêt à envoyer », « ⚠ Plan manquant », etc.).
   - Vous pouvez supprimer un client de la liste si besoin, ou tout vider.
4. Quand tous les clients du soir sont ajoutés, cliquez sur
   **"📧 Envoyer tous les emails"**, puis confirmez.
5. L'application envoie les emails un par un via Outlook et affiche la
   progression dans le journal en bas de l'écran.

L'objet généré automatiquement est :
`Votre Arrivé tardive - Camping Huttopia Lac de Rillé - Nom Prénom`.

### Option de vérification avant envoi

La case à cocher *"Ouvrir chaque email dans Outlook pour vérification avant
envoi"* permet, si vous préférez, d'ouvrir chaque email en brouillon dans
Outlook (vous cliquez vous-même sur Envoyer) plutôt que de tout envoyer
automatiquement. Pratique pour se rassurer les premiers jours d'utilisation.

---

## 5. Historique des envois

Chaque tentative d'envoi (réussie ou en échec) est enregistrée dans un fichier
`historique_envois.csv`, créé automatiquement à côté du `.exe`. Il s'ouvre
avec Excel et contient : date/heure, nom du client, emplacement(s), email,
statut, plans manquants éventuels, et le détail d'une erreur le cas échéant.

Accessible aussi directement depuis le menu **Fichier → Ouvrir l'historique
des envois** de l'application.

---

## 6. En cas de problème

- **"pywin32 n'est pas installé" / erreur Outlook au démarrage** → vérifiez
  qu'Outlook est bien installé et configuré sur le PC, et relancez
  `build.bat` pour être sûr que la dépendance `pywin32` est bien présente.
- **Un email part avec le mauvais plan / sans plan** → vérifiez que le nom du
  fichier PDF ou image correspond exactement au numéro d'emplacement saisi dans
  l'application (ex : emplacement "12" → fichier `12.pdf`, `12.png` ou
  `12.jpg`).
- **L'email part depuis la mauvaise boîte mail** → Outlook envoie depuis le
  compte par défaut configuré sur le PC. Vérifiez dans Outlook
  (Fichier → Informations) que le compte du camping est bien le compte par
  défaut.

---

## 7. Prochaines évolutions prévues (phase 2, pas encore codées)

- Import direct d'un fichier Excel contenant la liste des clients
  (Nom/Prénom + emplacement(s)), avec sélection des clients à qui envoyer le
  mail et saisie de leur adresse email une par une.
- Possibilité de joindre manuellement un plan personnalisé pour les clients
  ayant plusieurs hébergements.
