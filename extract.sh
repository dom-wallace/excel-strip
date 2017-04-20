#!/usr/local/bin/zsh

## A simple tool to strip the first attachement from an Excel file.

if [[ $# != 2 ]]; then
	echo "Usage: $0 SOURCE_DIRECTORY TARGET_DIRECTORY"
	exit 1
fi

FILES=$1
TARGET=$2
COUNT=0

## Bail out if the source or destination are bad
if [[ ! -d "$FILES" ]]; then
	echo "Source directory $FILES does not exist"
	exit 1
fi

if [[ ! -d "$TARGET" ]]; then
        echo "Target directory $TARGET does not exist"
        exit 1
fi

##Create a temp directory to work in
WORKDIR=$(mktemp -d)

if [[ ! "$WORKDIR" || ! -d "$WORKDIR" ]]; then
	echo "could not create $WORKDIR"
	exit 1
fi

## Cleanup function, then trap the EXIT signal
function cleanup {      
  rm -rf "$WORKDIR"
  echo "Deleted temp working directory $WORKDIR"
}

trap cleanup EXIT

## Loop through all the Excel files in the directory, and extract the first attachment.
for f in  $FILES/**/*.xlsx ; do
	unzip -q -d "$WORKDIR/$COUNT" "$f" 
	if [[ -f "$WORKDIR/$COUNT/xl/embeddings/oleObject1.bin" ]]; then
		echo "moving $WORKDIR/$COUNT/xl/embeddings/oleObject1.bin to $TARGET/invoice_$COUNT.pdf"
		mv "$WORKDIR/$COUNT/xl/embeddings/oleObject1.bin" "$TARGET/invoice_$COUNT.pdf" || { echo "could not move file to $TARGET"; exit 1; }
	else
		echo "no PDF found in $f. Skipping"
	fi	
	((COUNT++))
done

## We're done!
exit 0
