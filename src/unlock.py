import os
import re
import sys
import shutil
import zipfile
import posixpath
import xml.etree.ElementTree as ET
from pathlib import Path


# Regex patterns matching Excel protection tags (both self-closing and paired tags)
SHEET_PROTECTION_PATTERNS = [
    re.compile(r"<sheetProtection[^>]*(\/>|>.*?<\/sheetProtection>)", re.IGNORECASE | re.DOTALL),
    re.compile(r"<protectedRanges[^>]*(\/>|>.*?<\/protectedRanges>)", re.IGNORECASE | re.DOTALL),
    re.compile(r"<dialogsheetProtection[^>]*(\/>|>.*?<\/dialogsheetProtection>)", re.IGNORECASE | re.DOTALL),
]

WORKBOOK_STRUCTURE_PATTERNS = [
    re.compile(r"<workbookProtection[^>]*(\/>|>.*?<\/workbookProtection>)", re.IGNORECASE | re.DOTALL),
    re.compile(r'\s+lockStructure=["\']?(?:true|1)["\']?', re.IGNORECASE),
    re.compile(r'\s+lockWindows=["\']?(?:true|1)["\']?', re.IGNORECASE),
    re.compile(r'\s+lockRevision=["\']?(?:true|1)["\']?', re.IGNORECASE),
]

FILE_SHARING_PATTERNS = [
    re.compile(r"<fileSharing[^>]*(\/>|>.*?<\/fileSharing>)", re.IGNORECASE | re.DOTALL),
]


def inspect_excel_file(source: Path) -> list[dict]:
    """
    Analyse les éléments protégés d'un classeur Excel sans le modifier.
    Retourne la liste des éléments protégés avec leurs vrais noms de feuilles.
    """
    elements = []

    with zipfile.ZipFile(source, "r") as z:
        # 1. Protection du classeur
        has_wb_prot = False
        has_file_sharing = False

        if "xl/workbook.xml" in z.namelist():
            wb_xml = z.read("xl/workbook.xml").decode("utf-8", errors="ignore")
            has_wb_prot = any(p.search(wb_xml) for p in WORKBOOK_STRUCTURE_PATTERNS)
            has_file_sharing = any(p.search(wb_xml) for p in FILE_SHARING_PATTERNS)

        if has_wb_prot:
            elements.append({
                "id": "wb_structure",
                "type": "workbook",
                "title": "Structure du Classeur (Structure / Fenêtres)",
                "target": "xl/workbook.xml",
                "is_protected": True
            })

        if has_file_sharing:
            elements.append({
                "id": "wb_sharing",
                "type": "sharing",
                "title": "Partage / Mot de passe de modification",
                "target": "xl/workbook.xml",
                "is_protected": True
            })

        # 2. Relations des feuilles
        rel_map = {}
        if "xl/_rels/workbook.xml.rels" in z.namelist():
            try:
                root_rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
                for elem in root_rels:
                    r_id = elem.attrib.get("Id")
                    target = elem.attrib.get("Target", "").lstrip("/")
                    if not target.startswith("xl/"):
                        target = "xl/" + target
                    rel_map[r_id] = target
            except Exception:
                pass

        # 3. Liste des feuilles dans workbook.xml
        if "xl/workbook.xml" in z.namelist():
            try:
                root_wb = ET.fromstring(z.read("xl/workbook.xml"))
                sheet_idx = 1
                for elem in root_wb.iter():
                    if elem.tag.endswith("sheet") and "name" in elem.attrib:
                        s_name = elem.attrib.get("name")
                        s_rid = None
                        for k, v in elem.attrib.items():
                            if k.endswith("id"):
                                s_rid = v
                                break
                        
                        target_path = rel_map.get(s_rid, f"xl/worksheets/sheet{sheet_idx}.xml")
                        sheet_idx += 1

                        if target_path in z.namelist():
                            sheet_xml = z.read(target_path).decode("utf-8", errors="ignore")
                            is_prot = any(p.search(sheet_xml) for p in SHEET_PROTECTION_PATTERNS)
                            elements.append({
                                "id": f"sheet_{sheet_idx}",
                                "type": "sheet",
                                "title": f"Feuille « {s_name} »",
                                "target": target_path,
                                "is_protected": is_prot
                            })
            except Exception:
                pass

    return elements


def remove_excel_protection(xlsx_file: str | Path, targets_to_unlock: list[str] = None, make_backup: bool = False) -> bool:
    """
    Supprime les protections spécifiées (ou toutes par défaut) et remplace
    le fichier en place tout en conservant le même nom.

    :param xlsx_file: Chemin vers le fichier Excel.
    :param targets_to_unlock: Liste de chemins XML ou IDs à déverrouiller. Si None, déverrouille TOUT.
    :param make_backup: Si True, crée une sauvegarde '.bak'.
    :return: True si succès.
    """
    source = Path(xlsx_file).resolve()

    if not source.exists() or not source.is_file() or not zipfile.is_zipfile(source):
        print(f"[ERREUR] '{source.name}' n'est pas un fichier Excel valide.")
        return False

    if make_backup:
        backup_path = source.with_suffix(source.suffix + ".bak")
        shutil.copy2(source, backup_path)
        print(f"[INFO] Sauvegarde créée : {backup_path.name}")

    temp_file = source.parent / f".~{source.stem}_temp_unlock{source.suffix}"
    modifications = 0

    try:
        with zipfile.ZipFile(source, "r") as zin:
            with zipfile.ZipFile(temp_file, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)

                    if item.filename.lower().endswith((".xml", ".rels")):
                        try:
                            text = data.decode("utf-8")
                        except UnicodeDecodeError:
                            text = data.decode("latin-1", errors="ignore")

                        original_text = text

                        # 1. Classeur (Structure & Partage)
                        if item.filename == "xl/workbook.xml":
                            if targets_to_unlock is None or "wb_structure" in targets_to_unlock:
                                for p in WORKBOOK_STRUCTURE_PATTERNS:
                                    text, count = p.subn("", text)
                                    modifications += count
                            if targets_to_unlock is None or "wb_sharing" in targets_to_unlock:
                                for p in FILE_SHARING_PATTERNS:
                                    text, count = p.subn("", text)
                                    modifications += count

                        # 2. Feuilles de calcul
                        elif "worksheets/" in item.filename or item.filename.lower().endswith(".xml"):
                            should_unlock_sheet = True
                            if targets_to_unlock is not None:
                                should_unlock_sheet = item.filename in targets_to_unlock

                            if should_unlock_sheet:
                                for p in SHEET_PROTECTION_PATTERNS:
                                    text, count = p.subn("", text)
                                    modifications += count

                        if text != original_text:
                            data = text.encode("utf-8")

                    zout.writestr(item, data)

        os.replace(temp_file, source)

        if modifications > 0:
            print(f"[SUCCÈS] '{source.name}' déverrouillé avec succès ({modifications} élément(s) modifié(s)).")
        else:
            print(f"[INFO] '{source.name}' traité (aucune modification requise).")

        return True

    except Exception as e:
        print(f"[ERREUR] Échec sur '{source.name}': {e}")
        if temp_file.exists():
            temp_file.unlink(missing_ok=True)
        return False


def main():
    args = [arg for arg in sys.argv[1:] if not arg.startswith("--")]
    flags = [arg for arg in sys.argv[1:] if arg.startswith("--")]

    make_backup = "--backup" in flags
    inspect_only = "--inspect" in flags

    target_file = None
    if args:
        target_file = Path(args[0])
    elif Path("mon_fichier.xlsx").exists():
        target_file = Path("mon_fichier.xlsx")
    else:
        found = [f for f in Path(".").glob("*.xls*") if not f.name.startswith((".", "~$"))]
        data_dir = Path("data")
        if data_dir.is_dir():
            found += [f for f in data_dir.glob("*.xls*") if not f.name.startswith((".", "~$"))]

        if found:
            print("Fichiers Excel détectés :")
            for idx, f in enumerate(found, 1):
                parent_info = f" ({f.parent})" if str(f.parent) != "." else ""
                print(f"  [{idx}] {f.name}{parent_info}")
            choice = input("\nEntrez le numéro du fichier à traiter (ou 'q' pour quitter) : ").strip()
            if choice.isdigit() and 1 <= int(choice) <= len(found):
                target_file = found[int(choice) - 1]

    if not target_file or not target_file.exists():
        print("Aucun fichier valide sélectionné.")
        return

    print(f"\n--- Analyse des protections de '{target_file.name}' ---")
    elements = inspect_excel_file(target_file)

    if not elements:
        print("Aucun élément spécifique trouvé dans le classeur.")
        remove_excel_protection(target_file, make_backup=make_backup)
        return

    protected_elements = [e for e in elements if e["is_protected"]]
    print(f"Total : {len(protected_elements)} élément(s) protégé(s) détecté(s) :\n")

    for idx, e in enumerate(elements, 1):
        status = "[PROTÉGÉ]" if e["is_protected"] else "[Libre]"
        print(f"  [{idx}] {status} {e['title']}")

    if inspect_only:
        return

    print("\nOptions de déverrouillage :")
    print("  [Entrée]  : Déverrouiller TOUS les éléments protégés par défaut")
    print("  [1,2,...] : Déverrouiller uniquement certains numéros spécifiques (ex: 1,3)")
    print("  [q]       : Quitter")

    choice = input("\nVotre choix : ").strip()

    if choice.lower() == "q":
        print("Opération annulée.")
        return
    elif choice == "" or choice.lower() == "all":
        remove_excel_protection(target_file, targets_to_unlock=None, make_backup=make_backup)
    else:
        selected_targets = []
        indices = [c.strip() for c in choice.split(",") if c.strip().isdigit()]
        for idx_str in indices:
            i = int(idx_str) - 1
            if 0 <= i < len(elements):
                elem = elements[i]
                if elem["type"] in ("workbook", "sharing"):
                    selected_targets.append(elem["id"])
                else:
                    selected_targets.append(elem["target"])
        
        remove_excel_protection(target_file, targets_to_unlock=selected_targets, make_backup=make_backup)


if __name__ == "__main__":
    main()
