#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
output_dir="${repo_root}/output"
truth_dir="${repo_root}/truth"

if [[ ! -d "${output_dir}" ]]; then
  echo "output directory not found at: ${output_dir}" >&2
  exit 1
fi

updated=0

while IFS= read -r -d '' test_file; do
  relative_file="${test_file#"${output_dir}/"}"
  expected_file="${truth_dir}/${relative_file/.test./.}"
  if [[ "${expected_file}" == "${truth_dir}/${relative_file}" ]] || [[ ! -f "${expected_file}" ]]; then
    continue
  fi
  cp -f "${test_file}" "${expected_file}"
  updated=$((updated + 1))
done < <(find "${output_dir}" -type f -name "*.test.*" -print0)

echo "Updated ${updated} expected file(s)."
