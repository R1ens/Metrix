# -*- coding: utf-8 -*-
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path
from datetime import datetime


# ==========================
# ПУТИ (ничего не надо передавать)
# ==========================
ABS_INPUT_DIR = None   # например r"F:\Knowledge_bases\Bot_AFK_doc\PythonApplication1"
ABS_OUTPUT_DIR = None  # например r"F:\Knowledge_bases\Bot_AFK_doc\PythonApplication1\md_out"

SKIP_DIRS = {".vs", "md_out", "__pycache__", ".git"}
MEDIA_SUFFIX = "_media"

# Главная идея: сначала пробуем "чистые" writer'ы без html/div/span,
# но если pandoc не принимает расширения (из-за версии) — откатываемся к более простым.
TO_CANDIDATES = [
    # Самый "чистый" вариант: без raw_html (чтобы не лез HTML), и без дивов/спанов
    # (НЕ используем native_divs/native_spans — они часто ломаются на разных версиях).
    "gfm-raw_html-fenced_divs-bracketed_spans",

    # Чуть проще
    "gfm-raw_html-fenced_divs",
    "gfm-raw_html",

    # Просто gfm (если чистые режимы не поддерживаются конкретной версией)
    "gfm",

    # Альтернативный writer markdown
    "markdown-raw_html",
    "markdown",

    # Крайний fallback: plain (мы всё равно запишем как .md)
    "plain",
]

# Общие аргументы pandoc для читабельности
PANDOC_ARGS = [
    "--wrap=none",  # не ломать строки внутри абзацев
]


def check_pandoc() -> str:
    """Возвращает версию pandoc или завершает программу."""
    try:
        p = subprocess.run(["pandoc", "--version"], check=True, capture_output=True, text=True)
        first_line = (p.stdout.splitlines()[0] if p.stdout else "").strip()
        return first_line or "pandoc (version unknown)"
    except FileNotFoundError:
        print("ОШИБКА: pandoc не найден. Установи pandoc и перезапусти Visual Studio.", file=sys.stderr)
        raise SystemExit(1)
    except subprocess.CalledProcessError as e:
        print(f"ОШИБКА: pandoc запускается с ошибкой: {e}", file=sys.stderr)
        raise SystemExit(1)


def is_in_skipped_dir(path: Path, base: Path) -> bool:
    try:
        rel_parts = path.relative_to(base).parts
    except ValueError:
        return False
    return any(part in SKIP_DIRS for part in rel_parts)


# ---------- Очистка Markdown ----------
RE_FENCED_DIV_LINE = re.compile(r"^\s*:::+\s*(\{.*\})?\s*$")
RE_ATTR_TRAILING = re.compile(r"\s*\{[^}]*\}\s*$")
RE_HTML_TAG = re.compile(r"</?([A-Za-z][A-Za-z0-9]*)\b[^>]*>")
RE_MULTI_BLANK = re.compile(r"\n{3,}")


def clean_markdown(md_text: str) -> str:
    lines = md_text.splitlines()
    out_lines: list[str] = []
    for line in lines:
        if RE_FENCED_DIV_LINE.match(line):
            continue
        line = RE_ATTR_TRAILING.sub("", line)
        line = RE_HTML_TAG.sub("", line)
        out_lines.append(line.rstrip())

    text = "\n".join(out_lines)
    text = RE_MULTI_BLANK.sub("\n\n", text)
    text = "\n".join("" if l.strip() == "" else l for l in text.splitlines())
    return text.strip() + "\n"


def run_pandoc(docx_path: Path, out_md: Path, media_dir: Path, to_format: str) -> tuple[bool, str]:
    cmd = [
        "pandoc",
        "-f", "docx",
        "-t", to_format,
        f"--extract-media={str(media_dir)}",
        *PANDOC_ARGS,
        str(docx_path),
        "-o", str(out_md),
    ]
    try:
        p = subprocess.run(cmd, check=True, capture_output=True, text=True)
        warn = (p.stderr or "").strip()
        return True, warn
    except subprocess.CalledProcessError as e:
        msg = (e.stderr or e.stdout or "").strip()
        return False, msg


def convert_with_fallback(docx_path: Path, out_md: Path, media_dir: Path) -> tuple[bool, str, str]:
    """
    Пробует несколько -t форматов.
    Возвращает: (успех, выбранный_формат, сообщение)
    """
    out_md.parent.mkdir(parents=True, exist_ok=True)
    media_dir.mkdir(parents=True, exist_ok=True)

    last_error = ""
    for to_format in TO_CANDIDATES:
        ok, msg = run_pandoc(docx_path, out_md, media_dir, to_format)
        if ok:
            # Пост-очистка md
            try:
                md_text = out_md.read_text(encoding="utf-8")
            except UnicodeDecodeError:
                md_text = out_md.read_text(encoding="utf-8", errors="replace")
            out_md.write_text(clean_markdown(md_text), encoding="utf-8")
            return True, to_format, (msg or "")
        else:
            last_error = f"[{to_format}] {msg}"

    return False, "", last_error


def main() -> None:
    pandoc_ver = check_pandoc()

    script_dir = Path(__file__).resolve().parent
    input_dir = Path(ABS_INPUT_DIR).resolve() if ABS_INPUT_DIR else script_dir
    output_dir = Path(ABS_OUTPUT_DIR).resolve() if ABS_OUTPUT_DIR else (input_dir / "md_out")
    output_dir.mkdir(parents=True, exist_ok=True)

    log_path = output_dir / "_convert_log.txt"

    print("=== DOCX -> Clean Markdown batch converter (fallback) ===")
    print(f"PANDOC   : {pandoc_ver}")
    print(f"INPUT_DIR: {input_dir}")
    print(f"OUT_DIR  : {output_dir}")
    print("--------------------------------------------------------")

    docx_files: list[Path] = []
    for f in input_dir.rglob("*.docx"):
        if is_in_skipped_dir(f, input_dir):
            continue
        if f.name.startswith("~$"):
            continue
        docx_files.append(f)
    docx_files.sort()

    if not docx_files:
        print("Не найдено ни одного .docx в:", input_dir)
        return

    ok_count = 0
    fail_count = 0

    with log_path.open("a", encoding="utf-8") as log:
        log.write("\n" + "=" * 80 + "\n")
        log.write(f"RUN {datetime.now().isoformat(timespec='seconds')}\n")
        log.write(f"PANDOC: {pandoc_ver}\n")
        log.write(f"INPUT : {input_dir}\nOUTPUT: {output_dir}\nFILES : {len(docx_files)}\n")
        log.write("TO_CANDIDATES:\n  - " + "\n  - ".join(TO_CANDIDATES) + "\n")
        log.write("=" * 80 + "\n")

        for docx in docx_files:
            rel = docx.relative_to(input_dir)
            out_md = (output_dir / rel).with_suffix(".md")
            media_dir = out_md.parent / f"{out_md.stem}{MEDIA_SUFFIX}"

            success, used_to, msg = convert_with_fallback(docx, out_md, media_dir)

            if success:
                ok_count += 1
                print(f"[OK]   {rel}  ->  {used_to}")
                if msg:
                    log.write(f"[WARN] {rel} ({used_to}) :: {msg}\n")
            else:
                fail_count += 1
                print(f"[FAIL] {rel}")
                log.write(f"[FAIL] {rel}\n{msg}\n\n")

        log.write(f"RESULT: ok={ok_count}, fail={fail_count}\n")

    print("\nГотово.")
    print(f"Успешно: {ok_count}")
    print(f"С ошибками: {fail_count}")
    print(f"Лог: {log_path}")


if __name__ == "__main__":
    main()
