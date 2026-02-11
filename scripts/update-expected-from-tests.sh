#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
samples_dir="${repo_root}/samples"

if [[ ! -d "${samples_dir}" ]]; then
  echo "samples directory not found at: ${samples_dir}" >&2
  exit 1
fi

updated=0

while IFS= read -r -d '' test_file; do
  expected_file="${test_file/.test./.}"
  if [[ "${expected_file}" == "${test_file}" ]]; then
    continue
  fi
  cp -f "${test_file}" "${expected_file}"
  updated=$((updated + 1))
done < <(find "${samples_dir}" -type f -name "*.test.*" -print0)

echo "Updated ${updated} expected file(s)."
