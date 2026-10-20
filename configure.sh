#!/bin/sh
# Points manifest.xml at your hosting location.
# Usage:  ./configure.sh YOUR-GITHUB-USERNAME.github.io
# (The repo must be named "jp-signature-addin" so the paths line up.)
set -e
if [ -z "$1" ]; then
  echo "Usage: ./configure.sh <host>   e.g. ./configure.sh jpgravitt.github.io"
  exit 1
fi
HOST=$(printf '%s' "$1" | sed -e 's#^https\{0,1\}://##' -e 's#/*$##')
sed -i.bak "s#YOUR-HOST#${HOST}#g" manifest.xml && rm -f manifest.xml.bak
echo "manifest.xml now points to https://${HOST}/jp-signature-addin/"
