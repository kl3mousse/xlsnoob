#!/bin/bash

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
VERSION="${1:-$(node -p "require('${ROOT_DIR}/package.json').version")}"
VERSION="${VERSION#v}"
BUILD_DIR="${ROOT_DIR}/release/build"
PAYLOAD_DIR="${BUILD_DIR}/payload"
OUTPUT_DIR="${ROOT_DIR}/release"
PACKAGE_PATH="${OUTPUT_DIR}/xlsNoob-${VERSION}-macOS.pkg"

rm -rf "${BUILD_DIR}"
mkdir -p "${PAYLOAD_DIR}/Library/Application Support/xlsNoob" "${OUTPUT_DIR}"
cp "${ROOT_DIR}/manifest.xml" "${PAYLOAD_DIR}/Library/Application Support/xlsNoob/manifest.xml"

pkgbuild \
  --root "${PAYLOAD_DIR}" \
  --scripts "${ROOT_DIR}/installer/scripts" \
  --identifier "com.xlsnoob.excel-addin" \
  --version "${VERSION}" \
  --install-location "/" \
  "${PACKAGE_PATH}"

rm -rf "${BUILD_DIR}"
echo "Created ${PACKAGE_PATH}"
