#!/usr/bin/env python3
import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta, timezone
from email import policy
from email.parser import BytesParser
from pathlib import Path
from typing import Optional

from bs4 import BeautifulSoup
from dateutil import parser as dateparser
from pypdf import PdfReader, PdfWriter
from pypdf.xmp import XmpInformation


@dataclass
class ExtractedMetadata:
    title: Optional[str] = None
    author: Optional[str] = None
    published_date_iso: Optional[str] = None
    source_url: Optional[str] = None
    subject: Optional[str] = None
    keywords: Optional[str] = None
    language: Optional[str] = None
    publisher: Optional[str] = None
    archive_capture_iso: Optional[str] = None
    source_mime: Optional[str] = None
    content_sha256: Optional[str] = None
    filename: Optional[str] = None
    source_path: Optional[str] = None
    confidence: str = "derived"


def log(msg: str, log_path: Path) -> None:
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as f:
        f.write(line + "\n")
    print(line)


def resolve_browser(choice: str) -> Path:
    candidates = {
        "chrome": [
            Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
            Path(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
        ],
        "edge": [
            Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
            Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
        ],
    }
    if choice.lower() == "auto":
        order = ["chrome", "edge"]
    else:
        order = [choice.lower()]
    for browser in order:
        for path in candidates[browser]:
            if path.exists():
                return path
    raise RuntimeError("No supported browser found (Chrome/Edge).")


def parse_dt(dt_str: str) -> Optional[datetime]:
    if not dt_str:
        return None
    tzinfos = {
        "UTC": timezone.utc,
        "GMT": timezone.utc,
        "EST": timezone(timedelta(hours=-5)),
        "EDT": timezone(timedelta(hours=-4)),
        "CST": timezone(timedelta(hours=-6)),
        "CDT": timezone(timedelta(hours=-5)),
        "MST": timezone(timedelta(hours=-7)),
        "MDT": timezone(timedelta(hours=-6)),
        "PST": timezone(timedelta(hours=-8)),
        "PDT": timezone(timedelta(hours=-7)),
    }
    try:
        return dateparser.parse(dt_str, tzinfos=tzinfos)
    except Exception:
        return None


def clean_text(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    cleaned = re.sub(r"\s+", " ", value).strip()
    return cleaned or None


def first_nonempty(*values):
    for v in values:
        if isinstance(v, str):
            v = clean_text(v)
        if v:
            return v
    return None


def extract_json_ld(soup: BeautifulSoup) -> dict:
    out = {}
    scripts = soup.find_all("script", attrs={"type": re.compile(r"ld\+json", re.I)})
    for sc in scripts:
        text = sc.string or sc.get_text("", strip=True)
        if not text:
            continue
        try:
            payload = json.loads(text)
        except Exception:
            continue
        nodes = payload if isinstance(payload, list) else [payload]
        for n in nodes:
            if not isinstance(n, dict):
                continue
            if not out.get("headline") and isinstance(n.get("headline"), str):
                out["headline"] = n["headline"]
            if not out.get("datePublished") and isinstance(n.get("datePublished"), str):
                out["datePublished"] = n["datePublished"]
            if not out.get("url") and isinstance(n.get("url"), str):
                out["url"] = n["url"]
            if not out.get("publisher"):
                pub = n.get("publisher")
                if isinstance(pub, dict):
                    name = pub.get("name")
                    if isinstance(name, str):
                        out["publisher"] = name
            if not out.get("author"):
                auth = n.get("author")
                if isinstance(auth, dict) and isinstance(auth.get("name"), str):
                    out["author"] = auth.get("name")
                elif isinstance(auth, list):
                    names = []
                    for a in auth:
                        if isinstance(a, dict) and isinstance(a.get("name"), str):
                            names.append(a.get("name"))
                        elif isinstance(a, str):
                            names.append(a)
                    if names:
                        out["author"] = ", ".join(dict.fromkeys(names))
                elif isinstance(auth, str):
                    out["author"] = auth
    return out


def extract_from_mht(path: Path) -> ExtractedMetadata:
    raw = path.read_bytes()
    sha256 = hashlib.sha256(raw).hexdigest()
    msg = BytesParser(policy=policy.default).parsebytes(raw)

    top_headers = {k.lower(): str(v) for k, v in msg.items()}
    source_mime = msg.get_content_type().lower() if msg.get_content_type() else "application/octet-stream"

    header_url = first_nonempty(
        top_headers.get("snapshot-content-location"),
        top_headers.get("content-location"),
        top_headers.get("x-original-url"),
        top_headers.get("x-source-url"),
    )
    header_date = first_nonempty(
        top_headers.get("date"),
        top_headers.get("x-msfilelastmodified"),
    )

    html = None
    part_url = None
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type().lower()
            if ctype == "text/html" and html is None:
                try:
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset() or "utf-8"
                    html = payload.decode(charset, errors="replace") if payload else None
                except Exception:
                    html = None
                part_url = first_nonempty(part.get("Content-Location"), part_url)
    else:
        ctype = msg.get_content_type().lower()
        if ctype == "text/html":
            payload = msg.get_payload(decode=True)
            charset = msg.get_content_charset() or "utf-8"
            html = payload.decode(charset, errors="replace") if payload else None

    if not html:
        html = raw.decode("utf-8", errors="replace")

    soup = BeautifulSoup(html, "html.parser")

    def meta_content(attr, value):
        tag = soup.find("meta", attrs={attr: re.compile(rf"^{re.escape(value)}$", re.I)})
        return tag.get("content") if tag else None

    json_ld = extract_json_ld(soup)

    title = first_nonempty(
        meta_content("property", "og:title"),
        meta_content("name", "twitter:title"),
        meta_content("name", "title"),
        json_ld.get("headline"),
        soup.title.string if soup.title else None,
    )

    author = first_nonempty(
        meta_content("name", "author"),
        meta_content("property", "article:author"),
        meta_content("name", "parsely-author"),
        meta_content("name", "byline"),
        json_ld.get("author"),
    )

    published_raw = first_nonempty(
        meta_content("property", "article:published_time"),
        meta_content("name", "pubdate"),
        meta_content("name", "publishdate"),
        meta_content("name", "date"),
        meta_content("property", "og:published_time"),
        json_ld.get("datePublished"),
        header_date,
    )

    canonical = soup.find("link", attrs={"rel": re.compile(r"canonical", re.I)})
    canonical_href = canonical.get("href") if canonical else None

    source_url = first_nonempty(
        canonical_href,
        meta_content("property", "og:url"),
        json_ld.get("url"),
        part_url,
        header_url,
    )

    keywords = first_nonempty(
        meta_content("name", "keywords"),
        meta_content("property", "article:tag"),
    )

    publisher = first_nonempty(
        meta_content("property", "article:publisher"),
        meta_content("name", "publisher"),
        json_ld.get("publisher"),
    )

    language = None
    html_tag = soup.find("html")
    if html_tag:
        language = clean_text(html_tag.get("lang"))

    desc = first_nonempty(
        meta_content("name", "description"),
        meta_content("property", "og:description"),
    )

    dt = parse_dt(published_raw)
    published_date_iso = dt.astimezone(timezone.utc).isoformat() if dt else None

    return ExtractedMetadata(
        title=title,
        author=author,
        published_date_iso=published_date_iso,
        source_url=source_url,
        subject=desc,
        keywords=keywords,
        language=language,
        publisher=publisher,
        archive_capture_iso=datetime.fromtimestamp(path.stat().st_ctime, tz=timezone.utc).isoformat(),
        source_mime=source_mime,
        content_sha256=sha256,
    )


def filetime_fallback(path: Path) -> str:
    ts = path.stat().st_ctime
    dt = datetime.fromtimestamp(ts, tz=timezone.utc)
    return dt.isoformat()


def pdf_dt_from_iso(iso: str) -> str:
    dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
    dt_utc = dt.astimezone(timezone.utc)
    return dt_utc.strftime("D:%Y%m%d%H%M%S+00'00'")


def render_pdf(browser_exe: Path, in_file: Path, out_pdf: Path, profile_base: Path) -> subprocess.CompletedProcess:
    file_profile = profile_base / str(os.urandom(8).hex())
    file_profile.mkdir(parents=True, exist_ok=True)

    uri = in_file.resolve().as_uri()

    cmd = [
        str(browser_exe),
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        f"--user-data-dir={str(file_profile)}",
        f"--print-to-pdf={str(out_pdf)}",
        uri,
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)

    shutil.rmtree(file_profile, ignore_errors=True)
    return proc


def apply_pdf_metadata(pdf_path: Path, meta: ExtractedMetadata, source_file: Path, converted_iso: str) -> None:
    reader = PdfReader(str(pdf_path))
    writer = PdfWriter(clone_from=reader)

    published_iso = meta.published_date_iso or filetime_fallback(source_file)
    title = meta.title or source_file.stem

    subject_parts = []
    if meta.subject:
        subject_parts.append(meta.subject)
    if meta.source_url:
        subject_parts.append(f"Source URL: {meta.source_url}")
    if published_iso:
        subject_parts.append(f"Published: {published_iso}")
    subject = " | ".join(subject_parts)[:2000] if subject_parts else ""

    info = {
        "/Title": title,
        "/Author": meta.author or "Unknown",
        "/Subject": subject,
        "/Keywords": meta.keywords or "",
        "/Creator": "mht2pdf metadata pipeline",
        "/Producer": "mht2pdf + pypdf",
        "/CreationDate": pdf_dt_from_iso(published_iso),
        "/ModDate": pdf_dt_from_iso(converted_iso),
        "/SourceURL": meta.source_url or "",
        "/SourceFile": str(source_file),
        "/Publisher": meta.publisher or "",
        "/Language": meta.language or "",
        "/ArchiveCaptureDate": meta.archive_capture_iso or "",
        "/ContentSHA256": meta.content_sha256 or "",
        "/SourceMIME": meta.source_mime or "",
    }
    writer.add_metadata(info)

    # XMP packet with same high-value fields
    xmp = XmpInformation.create()
    xmp.dc_title = {"x-default": title}
    xmp.dc_creator = [meta.author] if meta.author else ["Unknown"]
    xmp.dc_identifier = [meta.source_url] if meta.source_url else []
    xmp.dc_description = {"x-default": subject}
    xmp.dc_date = [published_iso]
    if meta.keywords:
        xmp.dc_subject = [k.strip() for k in meta.keywords.split(",") if k.strip()]
    if meta.language:
        xmp.dc_language = [meta.language]
    if meta.publisher:
        xmp.dc_publisher = [meta.publisher]
    xmp.xmp_creatortool = "mht2pdf metadata pipeline"
    writer.xmp_metadata = xmp

    tmp_out = pdf_path.with_suffix(".tmp.pdf")
    with tmp_out.open("wb") as f:
        writer.write(f)
    tmp_out.replace(pdf_path)


def main() -> int:
    ap = argparse.ArgumentParser(description="Convert MHT/MHTML to PDF with embedded metadata")
    ap.add_argument("--source-root", required=True)
    ap.add_argument("--output-root")
    default_log_path = os.environ.get("MHT2PDF_LOG_PATH")
    ap.add_argument("--log-path", default=default_log_path)
    ap.add_argument("--browser", choices=["Auto", "Chrome", "Edge"], default="Auto")
    ap.add_argument("--max-files", type=int, default=0)
    ap.add_argument("--skip-existing", action="store_true")
    ap.add_argument("--recurse-subdirs", action="store_true")
    args = ap.parse_args()

    source_root = Path(args.source_root.strip()).resolve()
    has_explicit_output_root = bool(args.output_root)
    if args.output_root:
        output_root = Path(args.output_root.strip()).resolve()
    else:
        output_root = (source_root / "_pdf_archive").resolve()
    if args.log_path:
        log_path = Path(args.log_path.strip()).resolve()
    else:
        log_path = (output_root / "logs" / "convert.log").resolve()

    output_root.mkdir(parents=True, exist_ok=True)
    tmp_root = output_root.parent / "tmp"
    norm_root = tmp_root / "normalized-mhtml"
    profile_root = tmp_root / "chrome-profiles"
    norm_root.mkdir(parents=True, exist_ok=True)
    profile_root.mkdir(parents=True, exist_ok=True)

    browser_exe = resolve_browser(args.browser)
    log(f"Browser: {browser_exe}", log_path)
    log(f"Source: {source_root}", log_path)

    if args.recurse_subdirs:
        candidates = source_root.rglob("*")
    else:
        candidates = source_root.glob("*")
    files = [p for p in candidates if p.is_file() and p.suffix.lower() in {".mht", ".mhtml"}]
    files.sort()
    if args.max_files > 0:
        files = files[: args.max_files]
    log(f"Target count: {len(files)}", log_path)

    ok = 0
    fail = 0
    converted_iso = datetime.now(timezone.utc).isoformat()

    for src in files:
        rel = src.relative_to(source_root)
        if has_explicit_output_root:
            out_pdf = (output_root / rel).with_suffix(".pdf")
            log_rel = out_pdf.relative_to(output_root)
        elif args.recurse_subdirs:
            out_pdf = (src.parent / "_pdf_archive" / f"{src.stem}.pdf").resolve()
            log_rel = out_pdf.relative_to(source_root)
        else:
            out_pdf = (output_root / src.name).with_suffix(".pdf")
            log_rel = out_pdf.relative_to(output_root)
        out_pdf.parent.mkdir(parents=True, exist_ok=True)

        if args.skip_existing and out_pdf.exists() and out_pdf.stat().st_size > 0:
            log(f"SKIP existing: {log_rel}", log_path)
            continue

        render_input = src
        temp_norm = None

        try:
            meta = extract_from_mht(src)
            meta.filename = src.name
            meta.source_path = str(src)

            if src.suffix.lower() == ".mht":
                temp_norm = (norm_root / rel).with_suffix(".mhtml")
                temp_norm.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, temp_norm)
                render_input = temp_norm

            proc = render_pdf(browser_exe, render_input, out_pdf, profile_root)
            if not out_pdf.exists() or out_pdf.stat().st_size == 0:
                err = (proc.stderr or proc.stdout or "").strip().splitlines()[:1]
                log(f"FAIL render: {rel} :: {' '.join(err)}", log_path)
                fail += 1
                continue

            # Fallbacks requested by user
            if not meta.title:
                meta.title = src.stem
            if not meta.published_date_iso:
                meta.published_date_iso = filetime_fallback(src)

            apply_pdf_metadata(out_pdf, meta, src, converted_iso)

            # Sidecar for audit and future remapping
            sidecar = out_pdf.with_suffix(".metadata.json")
            sidecar.write_text(json.dumps(asdict(meta), indent=2, ensure_ascii=False), encoding="utf-8")

            log(f"OK: {log_rel}", log_path)
            ok += 1
        except Exception as e:
            fail += 1
            log(f"FAIL error: {rel} :: {e}", log_path)
        finally:
            if temp_norm and temp_norm.exists():
                temp_norm.unlink(missing_ok=True)

    log(f"DONE ok={ok} fail={fail}", log_path)
    return 2 if fail else 0


if __name__ == "__main__":
    sys.exit(main())
