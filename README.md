# excel-strip
A very simple tool for stripping out what is assumed to be an embedded PDF inside an Excel file. It will recursively look for all .xlsx files in a given source directory, and writing the extracted PDFs to the target.

## Requirements
* A relatively recent version of Zsh (Z Shell).
* Enough disk space for all the extracted files during processing.
* A Unix like environment (Linux, macOS, Cygwin, etc) with the usual tools like `unzip` and `mktemp`.

## Usage
`./extract.sh SOURCE_DIRECTORY TARGET_DIRECTORY`
For example:
`./extract.sh /tmp/input /tmp/output`

## Limitiations
There are heaps:
1. Only looks for files with a .xlsx extension.
2. Zero validation on whether what is extracted is a PDF, it just blindly renames it.
3. Only gets the first embedded attachment, so if there is more than one or the file that is selcted is the wrong one, too bad.
4. Only supports Excel. This could be ported easily to do Word, PowerPoint, etc with a few minor changes, such as the file extension and the embeddings directory.
5. All the destination PDFs are called 'invoice' as this was originally used to extract scanned invoices from finance spreadsheets and I haven't changed it yet.
