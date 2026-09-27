from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

VERSION = "0.2.0"

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    PdfReader = PdfWriter = None

try:
    from PIL import Image
except ImportError:
    Image = None

try:
    import fitz
except ImportError:
    fitz = None

try:
    from pdf2docx import Converter
except ImportError:
    Converter = None


class AppError(Exception):
    pass


def die(message: str):
    raise AppError(message)


def ok(message: str):
    print(f"[OK] {message}")


def find_soffice() -> str | None:
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for item in candidates:
        if item and Path(item).exists():
            return item
    return None


def merge_pdfs(files: list[str], output: str):
    if PdfReader is None:
        die("pypdf belum terinstall. Jalankan: python -m pip install -e .")
    writer = PdfWriter()
    for f in files:
        if not Path(f).is_file():
            die(f"File tidak ditemukan: {f}")
        reader = PdfReader(f)
        for page in reader.pages:
            writer.add_page(page)
    with open(output, "wb") as fh:
        writer.write(fh)
    ok(f"Merge selesai: {output}")


def split_pdf(input_file: str, output_dir: str):
    if PdfReader is None:
        die("pypdf belum terinstall.")
    if not Path(input_file).is_file():
        die(f"File tidak ditemukan: {input_file}")
    reader = PdfReader(input_file)
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    for i, page in enumerate(reader.pages, start=1):
        writer = PdfWriter()
        writer.add_page(page)
        target = out / f"page_{i:03d}.pdf"
        with target.open("wb") as fh:
            writer.write(fh)
    ok(f"{len(reader.pages)} halaman diexport ke: {out}")


def rotate_pdf(input_file: str, angle: int, output: str):
    if PdfReader is None:
        die("pypdf belum terinstall.")
    if not Path(input_file).is_file():
        die(f"File tidak ditemukan: {input_file}")
    if angle % 90 != 0:
        die("Angle harus kelipatan 90.")
    reader = PdfReader(input_file)
    writer = PdfWriter()
    for page in reader.pages:
        page.rotate(angle)
        writer.add_page(page)
    with open(output, "wb") as fh:
        writer.write(fh)
    ok(f"Rotate selesai: {output}")


def compress_pdf(input_file: str, output: str | None = None):
    if PdfReader is None:
        die("pypdf belum terinstall.")
    if not Path(input_file).is_file():
        die(f"File tidak ditemukan: {input_file}")
    src = Path(input_file)
    target = Path(output) if output else src.with_name(src.stem + "_compressed.pdf")
    reader = PdfReader(str(src))
    writer = PdfWriter()
    for page in reader.pages:
        try:
            page.compress_content_streams()
        except Exception:
            pass
        writer.add_page(page)
    with target.open("wb") as fh:
        writer.write(fh)
    ok(f"Compression selesai: {target}")
    print("[NOTE] Ini structural compression; PDF berisi gambar besar mungkin tidak banyak mengecil.")


def images_to_pdf(files: list[str], output: str):
    if Image is None:
        die("Pillow belum terinstall.")
    images = []
    try:
        for f in files:
            if not Path(f).is_file():
                die(f"File tidak ditemukan: {f}")
            img = Image.open(f).convert("RGB")
            images.append(img)
        if not images:
            die("Tidak ada gambar.")
        images[0].save(output, save_all=True, append_images=images[1:])
        ok(f"Images → PDF selesai: {output}")
    finally:
        for img in images:
            img.close()


def pdf_to_images(input_file: str, output_dir: str):
    if fitz is None:
        die("PyMuPDF belum terinstall.")
    if not Path(input_file).is_file():
        die(f"File tidak ditemukan: {input_file}")
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(input_file)
    try:
        for i, page in enumerate(doc, start=1):
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            pix.save(str(out / f"page_{i:03d}.png"))
        ok(f"{len(doc)} halaman → PNG di: {out}")
    finally:
        doc.close()


def convert_document(input_file: str, output_dir: str | None = None):
    src = Path(input_file)
    if not src.is_file():
        die(f"File tidak ditemukan: {src}")

    ext = src.suffix.lower()
    supported = {
        ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
        ".odt", ".ods", ".odp", ".rtf", ".txt"
    }
    if ext not in supported:
        die(f"Format belum didukung untuk convert: {ext}")

    outdir = Path(output_dir) if output_dir else src.parent
    outdir.mkdir(parents=True, exist_ok=True)

    soffice = find_soffice()
    if soffice:
        cmd = [
            soffice, "--headless", "--convert-to", "pdf",
            "--outdir", str(outdir), str(src)
        ]
        result = subprocess.run(cmd, capture_output=True, text=True)
        target = outdir / (src.stem + ".pdf")
        if result.returncode == 0 and target.exists():
            ok(f"Convert selesai: {target}")
            return
        die((result.stderr or result.stdout or "LibreOffice gagal melakukan conversion.").strip())

    # Optional Microsoft Word automation if available.
    if ext in {".doc", ".docx", ".rtf"}:
        try:
            import win32com.client  # type: ignore
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            try:
                doc = word.Documents.Open(str(src.resolve()))
                target = outdir / (src.stem + ".pdf")
                doc.SaveAs(str(target.resolve()), FileFormat=17)
                doc.Close(False)
            finally:
                word.Quit()
            if target.exists():
                ok(f"Convert selesai: {target}")
                return
        except Exception:
            pass

    die(
        "Engine conversion belum tersedia. Install LibreOffice atau Microsoft Word, "
        "lalu jalankan ARCDVNpdf convert lagi."
    )


def info_pdf(input_file: str):
    if PdfReader is None:
        die("pypdf belum terinstall.")
    if not Path(input_file).is_file():
        die(f"File tidak ditemukan: {input_file}")
    reader = PdfReader(input_file)
    size = Path(input_file).stat().st_size
    print("\nARCDVNpdf PDF INFO")
    print("-" * 30)
    print(f"File   : {Path(input_file).resolve()}")
    print(f"Size   : {size / 1024:.1f} KB")
    print(f"Pages  : {len(reader.pages)}")
    print(f"Encrypted: {reader.is_encrypted}")


def pdf_to_word(input_file: str, output: str | None = None):
    if Converter is None:
        die("pdf2docx belum terinstall. Jalankan: python -m pip install -e .")
    if not Path(input_file).is_file():
        die(f"File tidak ditemukan: {input_file}")
    
    src = Path(input_file)
    target = Path(output) if output else src.with_suffix(".docx")
    
    print(f"[INFO] Memulai konversi {src.name} ke DOCX...")
    try:
        cv = Converter(str(src))
        cv.convert(str(target), start=0, end=None)
        cv.close()
        if target.exists():
            ok(f"Konversi selesai: {target}")
        else:
            die("Konversi gagal.")
    except Exception as e:
        die(f"Gagal melakukan konversi: {e}")


def doctor():
    print("\nARCDVNpdf Doctor")
    print("=" * 32)
    print(f"Version     : {VERSION}")
    print(f"Python      : {sys.version.split()[0]}")
    print(f"Executable  : {sys.executable}")
    print(f"ARCDVNpdf   : {Path(__file__).resolve()}")
    print(f"pypdf       : {'OK' if PdfReader else 'MISSING'}")
    print(f"Pillow      : {'OK' if Image else 'MISSING'}")
    print(f"PyMuPDF     : {'OK' if fitz else 'MISSING'}")
    print(f"pdf2docx    : {'OK' if Converter else 'MISSING'}")
    print(f"LibreOffice : {find_soffice() or 'Not found'}")
    try:
        import win32com.client  # noqa: F401
        print("MS Word COM : Available")
    except Exception:
        print("MS Word COM : Not available")
    print()


def build_parser():
    parser = argparse.ArgumentParser(
        prog="ARCDVNpdf",
        description="ARCDVNpdf - Local PDF and document utility."
    )
    parser.add_argument("--version", action="version", version=f"ARCDVNpdf {VERSION}")

    sub = parser.add_subparsers(dest="command")

    p = sub.add_parser("merge", help="Gabungkan PDF")
    p.add_argument("files", nargs="+")
    p.add_argument("-o", "--output", required=True)

    p = sub.add_parser("split", help="Pisahkan PDF menjadi per halaman")
    p.add_argument("file")
    p.add_argument("-o", "--output", default="pages")

    p = sub.add_parser("rotate", help="Putar semua halaman PDF")
    p.add_argument("file")
    p.add_argument("--angle", type=int, required=True)
    p.add_argument("-o", "--output", required=True)

    p = sub.add_parser("compress", help="Compress struktur PDF")
    p.add_argument("file")
    p.add_argument("-o", "--output")

    p = sub.add_parser("convert", help="Convert dokumen ke PDF")
    p.add_argument("file")
    p.add_argument("-o", "--output-dir")

    p = sub.add_parser("images", help="Gabungkan gambar menjadi PDF")
    p.add_argument("files", nargs="+")
    p.add_argument("-o", "--output", required=True)

    p = sub.add_parser("pdf-to-images", help="Convert PDF menjadi PNG")
    p.add_argument("file")
    p.add_argument("-o", "--output", default="images")

    p = sub.add_parser("info", help="Lihat informasi PDF")
    p.add_argument("file")

    p = sub.add_parser("pdf-to-word", help="Convert PDF ke Word (DOCX)")
    p.add_argument("file")
    p.add_argument("-o", "--output")

    sub.add_parser("doctor", help="Cek instalasi ARCDVNpdf")

    return parser


def ask_gui_file(title="Pilih File", filetypes=[("All Files", "*.*")]):
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    root.destroy()
    return path

def ask_gui_files(title="Pilih Beberapa File", filetypes=[("All Files", "*.*")]):
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    paths = filedialog.askopenfilenames(title=title, filetypes=filetypes)
    root.destroy()
    return list(paths)

def ask_gui_folder(title="Pilih Folder"):
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askdirectory(title=title)
    root.destroy()
    return path

def ask_gui_save(title="Simpan Sebagai", filetypes=[("All Files", "*.*")], defaultextension=""):
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.asksaveasfilename(title=title, filetypes=filetypes, defaultextension=defaultextension)
    root.destroy()
    return path

def clean_path(p: str) -> str:
    p = p.strip()
    if p.startswith('"') and p.endswith('"'):
        return p[1:-1]
    if p.startswith("'") and p.endswith("'"):
        return p[1:-1]
    return p


def interactive():
    while True:
        os.system("cls" if os.name == "nt" else "clear")
        print("\n" + "=" * 48)
        print("             ARCDVNpdf")
        print("      PDF & Document Utility CLI")
        print("=" * 48)
        print("[1] Merge PDF")
        print("[2] Split PDF")
        print("[3] Rotate PDF")
        print("[4] Compress PDF")
        print("[5] Convert Document")
        print("[6] Images ↔ PDF")
        print("[7] PDF Info")
        print("[8] PDF to Word (DOCX)")
        print("[9] Help")
        print("[0] Exit")

        choice = input("\nPilih menu: ").strip()

        if choice == "0":
            break

        try:
            if choice == "1":
                print("Masukkan path PDF satu per satu, atau kosongkan lalu Enter untuk buka GUI:")
                first_file = input("PDF 1 [Enter=GUI]: ")
                first_file = clean_path(first_file)
                if not first_file:
                    files = ask_gui_files(title="Pilih File PDF untuk di-Merge", filetypes=[("PDF", "*.pdf")])
                    if files:
                        print(f"-> {len(files)} file terpilih.")
                else:
                    files = [first_file]
                    count = 2
                    while True:
                        f = input(f"PDF {count} (kosongkan jika selesai): ")
                        f = clean_path(f)
                        if not f:
                            break
                        files.append(f)
                
                if files:
                    out = input("Output PDF [Enter=GUI]: ")
                    out = clean_path(out)
                    if not out:
                        out = ask_gui_save(title="Simpan PDF Merge", filetypes=[("PDF", "*.pdf")], defaultextension=".pdf")
                        if out:
                            print(f"-> Output: {out}")
                    if out:
                        merge_pdfs(files, out)
            elif choice == "2":
                file = input("PDF [Enter=GUI]: ")
                file = clean_path(file)
                if not file:
                    file = ask_gui_file(title="Pilih File PDF", filetypes=[("PDF", "*.pdf")])
                    if file:
                        print(f"-> Terpilih: {file}")
                if file:
                    out = input("Folder output [kosong=GUI, ketik 'pages' utk default]: ")
                    out = clean_path(out)
                    if not out:
                        out = ask_gui_folder(title="Pilih Folder Output")
                        if out:
                            print(f"-> Folder: {out}")
                    if not out:
                        out = "pages"
                    split_pdf(file, out)
            elif choice == "3":
                file = input("PDF [Enter=GUI]: ")
                file = clean_path(file)
                if not file:
                    file = ask_gui_file(title="Pilih File PDF", filetypes=[("PDF", "*.pdf")])
                    if file:
                        print(f"-> Terpilih: {file}")
                if file:
                    angle_str = input("Angle (90/180/270): ").strip()
                    try:
                        angle = int(angle_str)
                        out = input("Output PDF [Enter=GUI]: ")
                        out = clean_path(out)
                        if not out:
                            out = ask_gui_save(title="Simpan PDF Rotated", filetypes=[("PDF", "*.pdf")], defaultextension=".pdf")
                            if out:
                                print(f"-> Output: {out}")
                        if out:
                            rotate_pdf(file, angle, out)
                    except ValueError:
                        print("[ERROR] Angle harus berupa angka.")
            elif choice == "4":
                file = input("PDF [Enter=GUI]: ")
                file = clean_path(file)
                if not file:
                    file = ask_gui_file(title="Pilih File PDF", filetypes=[("PDF", "*.pdf")])
                    if file:
                        print(f"-> Terpilih: {file}")
                if file:
                    out = input("Output PDF [Enter=Otomatis/GUI]: ")
                    out = clean_path(out)
                    if not out:
                        ask_out = ask_gui_save(title="Simpan PDF Compressed (Batal = Otomatis)", filetypes=[("PDF", "*.pdf")], defaultextension=".pdf")
                        if ask_out:
                            out = ask_out
                            print(f"-> Output: {out}")
                        else:
                            out = None
                    compress_pdf(file, out)
            elif choice == "5":
                file = input("Dokumen (Word/Excel/dll) [Enter=GUI]: ")
                file = clean_path(file)
                if not file:
                    file = ask_gui_file(title="Pilih Dokumen", filetypes=[("Dokumen", "*.doc *.docx *.xls *.xlsx *.ppt *.pptx *.odt *.ods *.odp *.rtf *.txt"), ("Semua", "*.*")])
                    if file:
                        print(f"-> Terpilih: {file}")
                if file:
                    out = input("Output folder [Enter=Folder asal Dokumen/GUI]: ")
                    out = clean_path(out)
                    if not out:
                        ask_out = ask_gui_folder(title="Pilih Folder Output (Batal = Folder asal file)")
                        if ask_out:
                            out = ask_out
                            print(f"-> Folder: {out}")
                        else:
                            out = None
                    convert_document(file, out)
            elif choice == "6":
                mode = input("1 = Images→PDF, 2 = PDF→Images: ").strip()
                if mode == "1":
                    print("Masukkan path Gambar satu per satu, atau kosongkan lalu Enter untuk buka GUI:")
                    first_file = input("Gambar 1 [Enter=GUI]: ")
                    first_file = clean_path(first_file)
                    if not first_file:
                        files = ask_gui_files(title="Pilih Gambar", filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp"), ("Semua", "*.*")])
                        if files:
                            print(f"-> {len(files)} file terpilih.")
                    else:
                        files = [first_file]
                        count = 2
                        while True:
                            f = input(f"Gambar {count} (kosongkan jika selesai): ")
                            f = clean_path(f)
                            if not f:
                                break
                            files.append(f)
                    
                    if files:
                        out = input("Output PDF [Enter=GUI]: ")
                        out = clean_path(out)
                        if not out:
                            out = ask_gui_save(title="Simpan PDF", filetypes=[("PDF", "*.pdf")], defaultextension=".pdf")
                            if out:
                                print(f"-> Output: {out}")
                        if out:
                            images_to_pdf(files, out)
                elif mode == "2":
                    file = input("PDF [Enter=GUI]: ")
                    file = clean_path(file)
                    if not file:
                        file = ask_gui_file(title="Pilih File PDF", filetypes=[("PDF", "*.pdf")])
                        if file:
                            print(f"-> Terpilih: {file}")
                    if file:
                        out = input("Folder output [kosong=GUI, ketik 'images' utk default]: ")
                        out = clean_path(out)
                        if not out:
                            out = ask_gui_folder(title="Pilih Folder Output")
                            if out:
                                print(f"-> Folder: {out}")
                        if not out:
                            out = "images"
                        pdf_to_images(file, out)
                else:
                    print("[ERROR] Pilihan tidak valid.")
            elif choice == "7":
                file = input("PDF [Enter=GUI]: ")
                file = clean_path(file)
                if not file:
                    file = ask_gui_file(title="Pilih File PDF", filetypes=[("PDF", "*.pdf")])
                    if file:
                        print(f"-> Terpilih: {file}")
                if file:
                    info_pdf(file)
            elif choice == "8":
                file = input("PDF [Enter=GUI]: ")
                file = clean_path(file)
                if not file:
                    file = ask_gui_file(title="Pilih File PDF", filetypes=[("PDF", "*.pdf")])
                    if file:
                        print(f"-> Terpilih: {file}")
                if file:
                    out = input("Output DOCX [Enter=Otomatis/GUI]: ")
                    out = clean_path(out)
                    if not out:
                        ask_out = ask_gui_save(title="Simpan Word (DOCX) (Batal = Otomatis)", filetypes=[("Word Document", "*.docx")], defaultextension=".docx")
                        if ask_out:
                            out = ask_out
                            print(f"-> Output: {out}")
                        else:
                            out = None
                    pdf_to_word(file, out)
            elif choice == "9":
                build_parser().print_help()
            else:
                print("[ERROR] Pilihan tidak valid.")
        except AppError as e:
            print(f"\n[ERROR] {e}")
        except Exception as e:
            print(f"\n[ERROR] Terjadi kesalahan: {e}")
            
        input("\nTekan Enter untuk kembali ke menu...")


def main():
    if len(sys.argv) == 1:
        interactive()
        return

    parser = build_parser()
    args = parser.parse_args()

    try:
        if args.command == "merge":
            merge_pdfs(args.files, args.output)
        elif args.command == "split":
            split_pdf(args.file, args.output)
        elif args.command == "rotate":
            rotate_pdf(args.file, args.angle, args.output)
        elif args.command == "compress":
            compress_pdf(args.file, args.output)
        elif args.command == "convert":
            convert_document(args.file, args.output_dir)
        elif args.command == "images":
            images_to_pdf(args.files, args.output)
        elif args.command == "pdf-to-images":
            pdf_to_images(args.file, args.output)
        elif args.command == "info":
            info_pdf(args.file)
        elif args.command == "pdf-to-word":
            pdf_to_word(args.file, args.output)
        elif args.command == "doctor":
            doctor()
        else:
            parser.print_help()
    except AppError as e:
        print(f"\n[ERROR] {e}")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\nDibatalkan.")
        raise SystemExit(130)


if __name__ == "__main__":
    main()
