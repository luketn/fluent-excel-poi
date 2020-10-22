#!/bin/sh

# This is how to release a version - update the version number, e.g. ./release 0.0.1
VERSION=${1:?no version supplied}

echo "Deleting tag for version ${VERSION}..."

git tag -d "${VERSION}"
git push origin ":${VERSION}"
