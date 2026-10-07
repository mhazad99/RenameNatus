"""
Read Natus_filenames_for_RBD.xlsx and locate matching folders under SEARCH_ROOT.

For each row:
  - look up the folder named in column "Nom fichier"
  - copy into OUTPUT_ROOT / {Code}_{Date}/
  - fully anonymize + rename so folder and files share the same stem
    (same approach as xltek_reader/test/python/test_anonymize_header.py)
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Config — edit these paths / flags here (not via command line)
# ---------------------------------------------------------------------------
SEARCH_ROOT = r"\\athena\\Xltek_Recherche\\DbData"  # Study database to search
# Mother folder that will contain every anonymized/renamed study folder.
OUTPUT_ROOT = r"\\athena\\Xltek_Recherche\\Rename_test"
DO_RENAME = True  # True: anonymize + rename; False: search/report only
MAKE_COPY = True  # Always copy first; originals stay untouched

# anonymize_header flags — True = wipe / anonymize that field
ANON_FIRSTNAME = False
ANON_LASTNAME = False
ANON_MIDDLENAME = False
ANON_PATIENT_ID = False
ANON_PATIENT_GUID = False
ANON_BIRTH_DATE = False
ANON_SEX = False
ANON_HANDEDNESS = False
ANON_ADDRESS = False
ANON_TELEPHONE = False
ANON_HEIGHT = False
ANON_WEIGHT = False
ANON_CHART_NO = False
ANON_BILLING_ID = False
ANON_REFERRING_PHYSICIAN = False
ANON_WARD = False
ANON_STUDY_CREATOR = False
ANON_STUDY_NAME = False
ANON_STUDY_GUID = False
ANON_STUDY_NUMBERS = False
# rename_study / new_study_name / make_copy / copy_folder are set in rename_study()
# ---------------------------------------------------------------------------

# Prefer a local Windows wrapper (.pyd) copied into this folder.
# Fall back to the xltek_reader build tree if needed.
_SCRIPT_DIR = Path(__file__).resolve().parent
_XLTEK_RELEASE = Path(r"E:\CEAMS\InfoPathProject\xltek_reader\build\Release")
'''for candidate in (_SCRIPT_DIR, _XLTEK_RELEASE):
    if str(candidate) not in sys.path:
        sys.path.insert(0, str(candidate))'''

try:
    import XltekReader
except ImportError as exc:  # pragma: no cover - depends on local .pyd
    raise SystemExit(
        "Could not import XltekReader. Copy XltekReader.cp310-win_amd64.pyd "
        f"into {_SCRIPT_DIR} (or keep the build at {_XLTEK_RELEASE}).\n"
        f"Import error: {exc}"
    ) from exc

EXCEL_PATH = _SCRIPT_DIR / "Natus_filenames_for_RBD.xlsx"
COL_CODE = "Code"
COL_DATE = "Date PSG"
COL_NOM = "Nom fichier"


def configured_path(value: str) -> Path:
    """Convert a local or UNC path while tolerating repeated separators."""
    normalized = value.strip().replace("/", "\\")
    if normalized.startswith("\\\\"):
        normalized = "\\\\" + "\\".join(
            part for part in normalized[2:].split("\\") if part
        )
    return Path(normalized)


@dataclass
class RowMatch:
    code: str
    date_psg: datetime
    nom_fichier: str
    found_path: Path | None
    new_name: str
    eeg_path: Path | None = None
    renamed_path: Path | None = None


def load_excel(excel_path: Path) -> pd.DataFrame:
    df = pd.read_excel(excel_path)
    required = {COL_CODE, COL_DATE, COL_NOM}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing columns in Excel: {sorted(missing)}")
    return df.dropna(subset=[COL_NOM]).copy()


def build_new_study_name(code: str, date_psg: datetime) -> str:
    """
    New study stem = Code + Date PSG.

    Must stay compatible with XltekReader RenameStudyFiles:
    1-200 letters, digits, '_' or '-'.
    Folder name and file stems will both use this value.
    """
    if isinstance(date_psg, datetime):
        date_str = date_psg.strftime("%Y-%m-%d")
    else:
        date_str = pd.to_datetime(date_psg).strftime("%Y-%m-%d")
    return f"{code}_{date_str}"


def find_folder(search_root: Path, nom_fichier: str) -> Path | None:
    """
    Search the database (folder tree) for a directory whose name matches Nom fichier.

    Exact name match first; if none, fall back to matching the UUID suffix after the
    last underscore (common in Natus export folder names).
    """
    target = nom_fichier.strip()
    if not target:
        return None

    for path in search_root.rglob(target):
        if path.is_dir() and path.name == target:
            return path

    if "_" in target:
        suffix = target.rsplit("_", 1)[-1].lower()
        for path in search_root.rglob("*"):
            if path.is_dir() and "_" in path.name:
                if path.name.rsplit("_", 1)[-1].lower() == suffix:
                    return path

    return None


def find_eeg_file(study_folder: Path) -> Path | None:
    """Pick the .eeg that matches the study folder name, else the first .eeg found."""
    exact = study_folder / f"{study_folder.name}.eeg"
    if exact.is_file():
        return exact
    eeg_files = sorted(study_folder.glob("*.eeg"))
    return eeg_files[0] if eeg_files else None


def study_copy_destination(output_root: Path, new_study_name: str) -> Path:
    """
    Destination for one study copy:
      OUTPUT_ROOT / {Code}_{YYYY-MM-DD}/

    After rename_study, files inside match that folder name
    (e.g. R239_2018-09-26/R239_2018-09-26.eeg).
    """
    return output_root / new_study_name


def rename_study(
    study_folder: Path,
    new_study_name: str,
    make_copy: bool = True,
    copy_folder: str = "",
) -> Path:
    """
    Full anonymize + rename via XltekReader.anonymize_header.

    With make_copy=True the original study is left untouched. copy_folder must be
    the per-study destination directory (mother / new_study_name) so the folder
    name already matches the renamed file stems.
    """
    eeg_path = find_eeg_file(study_folder)
    if eeg_path is None:
        raise FileNotFoundError(f"No .eeg file found in study folder: {study_folder}")

    reader = XltekReader.XltekReader()
    if not reader.open_file(str(eeg_path)):
        raise RuntimeError(f"Could not open {eeg_path}: {reader.get_last_error()}")

    try:
        # Anonymize selected fields (see ANON_* config above) on a copy, then
        # rename files so their stem matches new_study_name (same as the folder name).
        success = reader.anonymize_header(
            firstname=ANON_FIRSTNAME,
            lastname=ANON_LASTNAME,
            middlename=ANON_MIDDLENAME,
            patient_id=ANON_PATIENT_ID,
            patient_guid=ANON_PATIENT_GUID,
            birth_date=ANON_BIRTH_DATE,
            sex=ANON_SEX,
            handedness=ANON_HANDEDNESS,
            address=ANON_ADDRESS,
            telephone=ANON_TELEPHONE,
            height=ANON_HEIGHT,
            weight=ANON_WEIGHT,
            chart_no=ANON_CHART_NO,
            billing_id=ANON_BILLING_ID,
            referring_physician=ANON_REFERRING_PHYSICIAN,
            ward=ANON_WARD,
            study_creator=ANON_STUDY_CREATOR,
            study_name=ANON_STUDY_NAME,
            study_guid=ANON_STUDY_GUID,
            study_numbers=ANON_STUDY_NUMBERS,
            rename_study=True,
            new_study_name=new_study_name,
            make_copy=make_copy,
            copy_folder=copy_folder,
        )
        if not success:
            raise RuntimeError(
                f"anonymize_header failed for {eeg_path}: {reader.get_last_error()}"
            )
        return Path(reader.get_eeg_path())
    finally:
        reader.close_file()


def process_rows(
    df: pd.DataFrame,
    search_root: Path,
    output_root: Path,
    do_rename: bool = False,
    make_copy: bool = True,
) -> list[RowMatch]:
    results: list[RowMatch] = []

    for _, row in df.iterrows():
        code = str(row[COL_CODE]).strip()
        date_psg = row[COL_DATE]
        nom = str(row[COL_NOM]).strip()
        new_name = build_new_study_name(code, date_psg)
        found = find_folder(search_root, nom)
        eeg_path = find_eeg_file(found) if found is not None else None
        dest_folder = study_copy_destination(output_root, new_name)

        match = RowMatch(
            code=code,
            date_psg=pd.to_datetime(date_psg).to_pydatetime(),
            nom_fichier=nom,
            found_path=found,
            new_name=new_name,
            eeg_path=eeg_path,
        )
        results.append(match)

        if found is None:
            print(f"[NOT FOUND] {nom}")
            continue

        print(f"[FOUND]     {nom}")
        print(f"            path : {found}")
        print(f"            eeg  : {eeg_path if eeg_path else '(missing)'}")
        print(f"            copy -> {dest_folder}")
        print(f"            files stem -> {new_name}")

        if do_rename:
            if dest_folder.exists():
                print(f"[SKIP]      destination already exists: {dest_folder}")
                continue
            renamed = rename_study(
                found,
                new_name,
                make_copy=make_copy,
                copy_folder=str(dest_folder),
            )
            match.renamed_path = renamed
            print(f"            done  : {renamed}")

    return results


def main() -> int:
    if not SEARCH_ROOT.strip():
        print(
            "Set SEARCH_ROOT at the top of this script to your study database folder.",
            file=sys.stderr,
        )
        return 1
    if not OUTPUT_ROOT.strip():
        print(
            "Set OUTPUT_ROOT at the top of this script to the mother folder "
            "for all renamed studies.",
            file=sys.stderr,
        )
        return 1

    excel_path = EXCEL_PATH.resolve()
    search_root = configured_path(SEARCH_ROOT)
    output_root = configured_path(OUTPUT_ROOT)

    if not excel_path.is_file():
        print(f"Excel file not found: {excel_path}", file=sys.stderr)
        return 1
    if not search_root.is_dir():
        print(f"Search root is not a directory: {search_root}", file=sys.stderr)
        return 1

    output_root.mkdir(parents=True, exist_ok=True)

    df = load_excel(excel_path)
    print(f"Loaded {len(df)} row(s) from {excel_path.name}")
    print(f"Searching under: {search_root}")
    print(f"Mother output folder: {output_root}")
    print(f"DO_RENAME={DO_RENAME}  MAKE_COPY={MAKE_COPY}")
    print(f"XltekReader module: {getattr(XltekReader, '__file__', XltekReader)}")
    print("-" * 60)

    results = process_rows(
        df,
        search_root,
        output_root,
        do_rename=DO_RENAME,
        make_copy=MAKE_COPY,
    )

    found_count = sum(1 for r in results if r.found_path is not None)
    done_count = sum(1 for r in results if r.renamed_path is not None)
    print("-" * 60)
    print(f"Done: {found_count}/{len(results)} found, {done_count} copied+renamed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
