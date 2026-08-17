# Excel Unlocker

Outil de suppression des protections de feuilles de calcul et de classeurs Microsoft Excel (.xlsx, .xlsm). 
Conçu pour fonctionner sans aucun fichier exécutable (.exe) et compatible avec les environnements d'entreprise aux politiques de sécurité strictes.

---

## 1. Structure du Dépôt

```text
xlsx_unlocker/
│
├── data/                             # Fichiers Excel d'exemples et de tests
│   ├── test_classeur_protege.xlsx
│   ├── test_feuille_protegee.xlsx
│   └── test_multi_feuilles_mixte.xlsx
│
├── src/                              # Code source des scripts
│   ├── __init__.py
│   ├── unlock.py                     # Script Python de déverrouillage
│   └── unlock.ps1                    # Script PowerShell natif Windows
│
├── tests/                            # Tests automatisés
│   ├── __init__.py
│   └── test_unlock.py                # Suite de tests unitaires (Pytest)
│
├── unlocker.html                     # Application Web autonome (0 installation, 0 dépendance)
├── unlock.bat                        # Lanceur Windows automatique (Drag & Drop)
│
├── DOCUMENTATION.md                  # Documentation technique détaillée et modèle OpenXML
├── README.md                         # Présentation générale du projet
├── requirements.txt                  # Dépendances Python pour tests et manipulation
└── .gitignore                        # Fichiers et dossiers exclus du suivi Git
```

---

## 2. Modes d'Utilisation

### Mode 1 : Interface Web Autonome (Recommandé)
Ouvrez directement le fichier `unlocker.html` dans n'importe quel navigateur Web (Edge, Chrome, Firefox).
- Glissez-déposez le fichier Excel à traiter.
- Inspectez les protections détectées (feuilles individuelles, structure globale).
- Cochez ou décochez les éléments à libérer.
- Cliquez sur « Déverrouiller uniquement la sélection » pour télécharger le fichier déverrouillé sous son nom exact.

### Mode 2 : Ligne de Commande Python
```powershell
# Déverrouillage d'un fichier spécifique
python src/unlock.py data/test_multi_feuilles_mixte.xlsx

# Mode interactif avec liste des fichiers disponibles
python src/unlock.py
```

### Mode 3 : Script PowerShell Natif Windows (Sans Python)
```powershell
powershell -ExecutionPolicy Bypass -File src/unlock.ps1 "data/test_multi_feuilles_mixte.xlsx"
```

### Mode 4 : Lanceur Windows (Glisser-Déposer)
Glissez-déposez un fichier `.xlsx` directement sur le fichier `unlock.bat` dans l'explorateur de fichiers Windows.

---

## 3. Exécution des Tests Automatisés

```powershell
.venv\Scripts\pytest.exe
```

---

## 4. Documentation Complète

Pour une analyse détaillée du fonctionnement interne, des balises XML ciblées et du diagramme de flux d'exécution, consultez le fichier `DOCUMENTATION.md`.
