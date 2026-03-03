#!/usr/bin/env python3
import argparse
import concurrent.futures as futures
import csv
import datetime as dt
import json
import os
import pathlib
import subprocess
import sys
from typing import Dict, Optional, Tuple

def utc_now_iso() -> str:
    return dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"

def find_soffice(explicit: Optional[str]) -> str:
    if explicit:
        p = pathlib.Path(explicit)
        if p.exists():
            return str(p)
        raise FileNotFoundError(f"--soffice path not found: {explicit}")

    env_path = os.environ.get("SOFFICE_PATH")
    if env_path and pathlib.Path(env_path).exists():
        return env_path

    if os.name == "nt":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for c in candidates:
            if pathlib.Path(c).exists():
                return c

    return "soffice"

def expected_pdf_path(pub: pathlib.Path) -> pathlib.Path:
    return pub.parent / f"{pub.stem}.pdf"

def run_convert_one(
    pub_path: str,
    soffice_path: str,
    timeout_sec: int,
    whatif: bool,
    skip_existing: bool
) -> Dict:
    pub = pathlib.Path(pub_path)
    outdir = pub.parent
    exp_pdf = expected_pdf_path(pub)

    result = {
        "timestamp_utc": utc_now_iso(),
        "pub_path": str(pub),
        "pdf_path": str(exp_pdf),
        "status": "fail",
        "error": "",
        "returncode": None,
        "whatif": bool(whatif),
    }

    if not pub.exists():
        result["error"] = "PUB file not found"
        return result

    if skip_existing and exp_pdf.exists() and exp_pdf.stat().st_size > 0:
        result["status"] = "skipped_existing"
        return result

    # WHATIF mode: simulate success (no conversion performed)
    if whatif:
        result["status"] = "success_simulated"
        result["returncode"] = 0
        result["error"] = ""
        return result

    cmd = [
        soffice_path,
        "--headless",
        "--nologo",
        "--nolockcheck",
        "--nodefault",
        "--norestore",
        "--convert-to", "pdf",
        "--outdir", str(outdir),
        str(pub),
    ]

    try:
        proc = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout_sec,
            text=True,
        )
        result["returncode"] = proc.returncode

        if proc.returncode != 0:
            result["error"] = (proc.stderr or proc.stdout or "").strip()[:2000]
            return result

        # verify expected output exists
        if exp_pdf.exists() and exp_pdf.stat().st_size > 0:
            result["status"] = "success"
            return result

        # fallback: sometimes LO produces a slightly different filename
        candidates = sorted(
            outdir.glob(pub.stem + "*.pdf"),
            key=lambda p: p.stat().st_mtime,
            reverse=True
        )
        if candidates and candidates[0].stat().st_size > 0:
            result["pdf_path"] = str(candidates[0])
            result["status"] = "success"
            return result

        result["error"] = "Conversion reported success but PDF not found or empty"
        return result

    except subprocess.TimeoutExpired:
        result["error"] = f"Timeout after {timeout_sec} seconds"
        return result
    except Exception as e:
        result["error"] = f"Exception: {type(e).__name__}: {e}"
        return result

def iter_pub_files(root: pathlib.Path) -> Tuple[str, ...]:
    return tuple(str(p) for p in root.rglob("*.pub"))

def write_csv(path: pathlib.Path, rows: list) -> None:
    fieldnames = ["timestamp_utc", "status", "pub_path", "pdf_path", "returncode", "whatif", "error"]
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in fieldnames})

def write_jsonl(path: pathlib.Path, rows: list) -> None:
    with path.open("w", encoding="utf-8") as f:
        for r in rows:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")

def main():
    ap = argparse.ArgumentParser(
        description="Batch convert Microsoft Publisher (.pub) files to PDF using LibreOffice. Outputs PDFs next to .pub files."
    )
    ap.add_argument("root", help="Root folder to scan recursively for .pub files")
    ap.add_argument("--soffice", help="Path to LibreOffice soffice executable (optional)")
    ap.add_argument("--workers", type=int, default=max(1, (os.cpu_count() or 4) - 1),
                    help="Parallel workers (default: CPU count minus 1)")
    ap.add_argument("--timeout", type=int, default=180, help="Per-file timeout seconds (default: 180)")
    ap.add_argument("--manifest-dir", default=None,
                    help="Where to write manifest files (default: root folder)")

    ap.add_argument("--whatif", action="store_true",
                    help="Simulate conversions without creating PDFs (no LibreOffice call).")
    ap.add_argument("--skip-existing", action="store_true",
                    help="Skip conversion if expected PDF already exists and is non-empty.")

    args = ap.parse_args()

    root = pathlib.Path(args.root).expanduser().resolve()
    if not root.exists() or not root.is_dir():
        print(f"Root folder not found or not a directory: {root}", file=sys.stderr)
        sys.exit(2)

    # Even in whatif mode, we still resolve soffice for parity (and to catch missing installs early)
    soffice_path = find_soffice(args.soffice)

    pubs = iter_pub_files(root)
    if not pubs:
        print(f"No .pub files found under: {root}")
        sys.exit(0)

    manifest_dir = pathlib.Path(args.manifest_dir).expanduser().resolve() if args.manifest_dir else root
    manifest_dir.mkdir(parents=True, exist_ok=True)

    csv_path = manifest_dir / "conversion_manifest.csv"
    jsonl_path = manifest_dir / "conversion_manifest.jsonl"

    print(f"Root: {root}")
    print(f"Found: {len(pubs)} .pub files")
    print(f"Using soffice: {soffice_path}")
    print(f"Workers: {args.workers}")
    print(f"WhatIf: {args.whatif}")
    print(f"SkipExisting: {args.skip_existing}")
    print(f"Manifest: {csv_path} and {jsonl_path}")
    print("Starting...")

    rows = []
    counts = {
        "success": 0,
        "success_simulated": 0,
        "skipped_existing": 0,
        "fail": 0,
    }

    with futures.ProcessPoolExecutor(max_workers=args.workers) as ex:
        tasks = [
            ex.submit(
                run_convert_one,
                p,
                soffice_path,
                args.timeout,
                args.whatif,
                args.skip_existing
            )
            for p in pubs
        ]

        total = len(tasks)
        for i, t in enumerate(futures.as_completed(tasks), start=1):
            r = t.result()
            rows.append(r)

            s = r.get("status", "fail")
            if s in counts:
                counts[s] += 1
            else:
                counts["fail"] += 1

            if i % 10 == 0 or i == total:
                print(
                    f"Progress: {i}/{total} "
                    f"(success={counts['success']}, "
                    f"simulated={counts['success_simulated']}, "
                    f"skipped={counts['skipped_existing']}, "
                    f"fail={counts['fail']})"
                )

    write_csv(csv_path, rows)
    write_jsonl(jsonl_path, rows)

    print("Done.")
    print(
        f"Totals: success={counts['success']}, "
        f"simulated={counts['success_simulated']}, "
        f"skipped_existing={counts['skipped_existing']}, "
        f"fail={counts['fail']}"
    )
    print(f"Manifest CSV:  {csv_path}")
    print(f"Manifest JSONL:{jsonl_path}")

if __name__ == "__main__":
    main()
