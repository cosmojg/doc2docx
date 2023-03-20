import json
import subprocess
import sys
from importlib.metadata import version
from pathlib import Path

from tqdm.auto import tqdm

__version__ = version(__package__)


def windows(paths, keep_active):
    import win32com.client

    word = win32com.client.Dispatch("Word.Application")
    wdFormatDocumentDefault = 16

    if paths["batch"]:
        for doc_filepath in tqdm(sorted(Path(paths["input"]).glob("[!~]*.doc*"))):
            docx_filepath = Path(paths["output"]) / f"{str(doc_filepath.stem)}.docx"
            doc = word.Documents.Open(str(doc_filepath))
            doc.SaveAs(str(docx_filepath), FileFormat=wdFormatDocumentDefault)
            doc.Close(0)
    else:
        pbar = tqdm(total=1)
        doc_filepath = Path(paths["input"]).resolve()
        docx_filepath = Path(paths["output"]).resolve()
        doc = word.Documents.Open(str(doc_filepath))
        doc.SaveAs(str(docx_filepath), FileFormat=wdFormatDocumentDefault)
        doc.Close(0)
        pbar.update(1)

    if not keep_active:
        word.Quit()


def macos(paths, keep_active):
    script = (Path(__file__).parent / "convert.jxa").resolve()
    cmd = [
        "/usr/bin/osascript",
        "-l",
        "JavaScript",
        str(script),
        str(paths["input"]),
        str(paths["output"]),
        str(keep_active).lower(),
    ]

    def run(cmd):
        process = subprocess.Popen(cmd, stderr=subprocess.PIPE)
        while True:
            line = process.stderr.readline().rstrip()
            if not line:
                break
            yield line.decode("utf-8")

    total = len(list(Path(paths["input"]).glob("*.doc*"))) if paths["batch"] else 1
    pbar = tqdm(total=total)
    for line in run(cmd):
        try:
            msg = json.loads(line)
        except ValueError:
            continue
        if msg["result"] == "success":
            pbar.update(1)
        elif msg["result"] == "error":
            print(msg)
            sys.exit(1)


def resolve_paths(input_path, output_path):
    input_path = Path(input_path).resolve()
    output_path = Path(output_path).resolve() if output_path else None
    output = {}
    if input_path.is_dir():
        output["batch"] = True
        output["input"] = str(input_path)
        if output_path:
            assert output_path.is_dir()
        else:
            output_path = str(input_path)
    else:
        output["batch"] = False
        assert str(input_path).endswith((".doc", ".DOC"))
        output["input"] = str(input_path)
        if output_path and output_path.is_dir():
            output_path = str(output_path / f"{str(input_path.stem)}.docx")
        elif output_path:
            assert str(output_path).endswith(".docx")
        else:
            output_path = str(input_path.parent / f"{str(input_path.stem)}.docx")
    output["output"] = output_path
    return output


def convert(input_path, output_path=None, keep_active=False):
    paths = resolve_paths(input_path, output_path)
    if sys.platform == "darwin":
        return macos(paths, keep_active)
    elif sys.platform == "win32":
        return windows(paths, keep_active)
    else:
        raise NotImplementedError(
            (
                "doc2docx is not implemented for linux as it requires Microsoft Word to"
                " be installed"
            ),
        )


def cli():
    import argparse
    import textwrap

    if "--version" in sys.argv:
        print(__version__)
        sys.exit(0)

    description = textwrap.dedent(
        """
    Example Usage:

    Convert single doc file in-place from myfile.doc to myfile.docx:
        doc2docx myfile.doc

    Batch convert doc folder in-place. Output docx files will go in the same folder:
        doc2docx myfolder/

    Convert single doc file with explicit output filepath:
        doc2docx input.doc output.doc

    Convert single doc file and output to a different explicit folder:
        doc2docx input.doc output_dir/

    Batch convert doc folder. Output docx files will go to a different explicit folder:
        doc2docx input_dir/ output_dir/
    """,
    )

    def formatter_class(prog):
        return argparse.RawDescriptionHelpFormatter(prog, max_help_position=32)

    parser = argparse.ArgumentParser(
        description=description,
        formatter_class=formatter_class,
    )
    parser.add_argument(
        "input",
        help=(
            "input file or folder. batch converts entire folder or convert single file"
        ),
    )
    parser.add_argument("output", nargs="?", help="output file or folder")
    parser.add_argument(
        "--keep-active",
        action="store_true",
        default=False,
        help="prevent closing word after conversion",
    )
    parser.add_argument(
        "--version",
        action="store_true",
        default=False,
        help="display version and exit",
    )

    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(0)
    else:
        args = parser.parse_args()

    convert(args.input, args.output, args.keep_active)
