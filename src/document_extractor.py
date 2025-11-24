import argparse
import hashlib
import shelve
from pathlib import Path
from typing import Iterator

from extractor import DrawioExtractor, ExcelExtractor, ExtractorBase, WordExtractor

parser = argparse.ArgumentParser(
    description="各種ドキュメントからテキストを抽出するツール"
)
parser.add_argument("target", type=Path, help="処理対象のディレクトリパス")
parser.add_argument(
    "--output",
    "-o",
    type=Path,
    help="抽出したテキストの出力先ディレクトリ",
    default=Path(".docs"),
)
parser.add_argument(
    "--force",
    "-f",
    action="store_true",
    help="ファイルハッシュをチェックせずに強制的に再抽出する",
)
args = parser.parse_args()

EXTRACTORS: dict[str, type[ExtractorBase]] = {
    ".xlsx": ExcelExtractor,
    ".docx": WordExtractor,
    ".drawio": DrawioExtractor,
}


def get_file_hash(path: Path, algo: str = "md5", chunk_size: int = 8192) -> str:
    h = hashlib.new(algo)
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(chunk_size), b""):
            h.update(chunk)
    return h.hexdigest()


def open_shelf(path: Path, shelf_file_name: str = ".index") -> shelve.Shelf:
    shelf_path = path / shelf_file_name
    shelf_path.parent.mkdir(parents=True, exist_ok=True)
    return shelve.open(shelf_path, flag="c")


def iter_target_files(
    target_dir: Path, extractors: dict[str, type[ExtractorBase]] = EXTRACTORS
) -> Iterator[tuple[Path, type[ExtractorBase]]]:
    for org_p in target_dir.rglob("*"):
        if not org_p.is_file():
            continue
        ext = org_p.suffix.lower()
        extractor_cls = extractors.get(ext)
        if extractor_cls is None:
            continue
        yield org_p, extractor_cls


def get_relative_path(original_path: Path, target_dir: Path) -> Path:
    return original_path.relative_to(target_dir)


def get_text_path(relative_path: Path, output_dir: Path) -> Path:
    return (output_dir / relative_path).with_name(relative_path.name + ".txt")


def save_hash(db: shelve.Shelf, relative_path: Path, hash: str) -> None:
    db[relative_path.as_posix()] = hash


def save_extract_file(relative_path: Path, output_dir: Path, extract: str) -> None:
    extract_p = get_text_path(relative_path, output_dir)
    extract_p.parent.mkdir(parents=True, exist_ok=True)
    with open(extract_p, mode="w", encoding="utf-8") as f:
        f.write(extract)


def delete_hash(db: shelve.Shelf, key: str) -> None:
    del db[key]


def delete_extract_file(relative_path: Path, output_dir: Path) -> None:
    extract_p = get_text_path(relative_path, output_dir)
    if extract_p.exists():
        extract_p.unlink()


def main():
    target: Path = args.target
    output: Path = args.output
    force: bool = args.force

    existing_paths: set[str] = set()
    db = open_shelf(output)
    for org_p, extractor_cls in iter_target_files(target):
        relative_path = get_relative_path(org_p, target)
        existing_paths.add(relative_path.as_posix())
        # 抽出済み（ハッシュが一致）したらスキップ
        hash = get_file_hash(org_p)
        if not force and hash == db.get(relative_path.as_posix()):
            print(f"Skip {relative_path}")
            continue

        print(f"Extracting {relative_path}")
        save_hash(db, relative_path, hash)
        save_extract_file(relative_path, output, extractor_cls(org_p).extract())

    # 削除済みファイルの反映
    to_delete_paths = set(db.keys()) - existing_paths
    for del_path in to_delete_paths:
        print(f"Removing {del_path}")
        delete_hash(db, del_path)
        delete_extract_file(Path(del_path), output)


if __name__ == "__main__":
    main()
