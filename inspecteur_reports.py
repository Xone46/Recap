from __future__ import annotations

import argparse
from pathlib import Path

from stamp_reports import DEFAULT_SIGNATURE, ROOT, _safe_name, process_inspector_source


def main() -> None:
    parser = argparse.ArgumentParser(description="Ajoute uniquement la signature de l'Inspecteur Agree dans les rapports DOCX.")
    parser.add_argument("source", help="Dossier ou archive .zip/.rar contenant les rapports DOCX")
    parser.add_argument("-o", "--output", help="Dossier de sortie pour les rapports signes")
    parser.add_argument("--signature", help="Chemin de signature (par defaut: signature/amine_foura.png)")
    args = parser.parse_args()

    source = Path(args.source)
    output = Path(args.output) if args.output else None
    signature = Path(args.signature) if args.signature else None
    results = process_inspector_source(source, output, signature)
    out_dir = output.resolve() if output else (ROOT / "outputs" / f"{_safe_name(source.name)}_inspecteur").resolve()
    signed_count = sum(line.startswith("OK\t") for line in results)
    print(f"source={source}")
    print(f"output={out_dir}")
    print(f"signature={signature or DEFAULT_SIGNATURE}")
    print(f"processed={len(results)} signed={signed_count}")
    for result in results:
        print(result)


if __name__ == "__main__":
    main()
