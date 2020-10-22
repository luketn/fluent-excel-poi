#!/bin/sh

# This is how to release a version - update the version number, e.g. ./release 0.0.1
VERSION=${1:?no version supplied}

echo "Publishing tagged version ${VERSION}..."

git tag -a "${VERSION}" -m "Release ${VERSION}"
git push origin "${VERSION}"
