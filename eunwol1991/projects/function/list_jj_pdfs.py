from __future__ import annotations

from pathlib import Path


WINDOWS_TARGET = Path(r"C:\Users\jhunj\Dropbox\for jj\Doc to print - JJ")
WSL_TARGET = Path("/mnt/c/Users/jhunj/Dropbox/for jj/Doc to print - JJ")


def resolve_target() -> Path:
    if WINDOWS_TARGET.exists():
        return WINDOWS_TARGET
    if WSL_TARGET.exists():
        return WSL_TARGET
    return WINDOWS_TARGET


def main() -> int:
    target = resolve_target()
    if not target.exists() or not target.is_dir():
        print(f"Target folder does not exist: {target}")
        return 1

    pdf_files = sorted(
        (p for p in target.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf"),
        key=lambda p: str(p).lower(),
    )

    for pdf in pdf_files:
        print(str(pdf))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
