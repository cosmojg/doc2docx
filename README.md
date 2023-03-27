# doc2docx

[![PyPI](https://img.shields.io/pypi/v/doc2docx)](https://pypi.org/project/doc2docx/)

Convert `doc` to `docx` on Windows or macOS directly using Microsoft Word (**must be installed**).

On Windows, this is implemented via [`win32com`](https://pypi.org/project/pywin32/) while on macOS this is implemented via [JXA](https://github.com/JXA-Cookbook/JXA-Cookbook) (Javascript for Automation, aka AppleScript in JS).

## Install

Via brew:

```
brew install cosmojg/tap/doc2docx
```

Via [pipx](https://pipxproject.github.io/pipx/):

```
pipx install doc2docx
```

Via pip:

```
pip install doc2docx
```

## CLI

```
usage: doc2docx [-h] [--keep-active] [--version] input [output]

Example Usage:

Convert single doc file in-place from myfile.doc to myfile.docx:
    doc2docx myfile.doc

Batch convert doc folder in-place. Output docx files will go in the same folder:
    doc2docx myfolder/

Convert single doc file with explicit output filepath:
    doc2docx input.doc output.docx

Convert single doc file and output to a different explicit folder:
    doc2docx input.doc output_dir/

Batch convert doc folder. Output docx files will go to a different explicit folder:
    doc2docx input_dir/ output_dir/

positional arguments:
  input          input file or folder. batch converts entire folder or convert
                 single file
  output         output file or folder

optional arguments:
  -h, --help     show this help message and exit
  --keep-active  prevent closing word after conversion
  --version      display version and exit
```

## Library

```python
from doc2docx import convert

convert("input.doc")
convert("input.doc", "output.docx")
convert("my_doc_folder/")
```

See CLI docs above (or in `doc2docx --help`) for all the different invocations. It is the same for the CLI and python library.

## Jupyter Notebook

If you are using this in the context of jupyter notebook, you will need `ipywidgets` for the tqdm progress bar to render properly.

```
pip install ipywidgets
jupyter nbextension enable --py widgetsnbextension
```

## Acknowledgements

Many thanks to [@AlJohri](https://github.com/AlJohri) for the excellent
[docx2pdf](https://github.com/AlJohri/docx2pdf) upon which this is based!
