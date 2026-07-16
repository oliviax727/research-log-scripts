#!/usr/bin/env bash
set -euo pipefail

SRC_FILE="${1:-modules.ts}"
shift $(( $# > 0 ? 1 : 0 )) || true

if [[ ! -f "$SRC_FILE" ]]; then
  echo "Source file not found: $SRC_FILE" >&2
  exit 1
fi

if ! command -v jq >/dev/null 2>&1; then
  echo "This script requires jq." >&2
  exit 1
fi

# Targets can be provided explicitly; otherwise use existing .osts files.
if [[ $# -gt 0 ]]; then
  targets=("$@")
else
  shopt -s nullglob
  targets=( *.osts )
  shopt -u nullglob
fi

if [[ ${#targets[@]} -eq 0 ]]; then
  echo "No output .osts targets found. Pass target files as arguments or create at least one .osts file." >&2
  exit 1
fi

escape_sed_replacement() {
  printf '%s' "$1" | sed -e 's/[\\/&]/\\\\&/g'
}

default_parameter_info='{"version":1,"originalParameterOrder":[],"parameterSchema":{"type":"object","default":{},"x-ms-visibility":"internal"},"returnSchema":{"type":"object","properties":{}},"signature":{"comment":"","parameters":[{"name":"workbook","comment":""}]}}'
default_api_info='{"variant":"synchronous","variantVersion":2}'

for target in "${targets[@]}"; do
  if [[ "$target" != *.osts ]]; then
    target="${target}.osts"
  fi

  script_name="${target##*/}"
  script_name="${script_name%.osts}"
  escaped_name="$(escape_sed_replacement "$script_name")"

  body="$(sed -E "s/(ResearchLog\.runFunction\(workbook, \")[^\"]+(\"\);)/\\1${escaped_name}\\2/" "$SRC_FILE")"

  if [[ -f "$target" ]]; then
    version="$(jq -r '.version // "0.3.0"' "$target")"
    description="$(jq -r '.description // empty' "$target")"
    no_code_metadata="$(jq -r '.noCodeMetadata // ""' "$target")"
    parameter_info="$(jq -r --arg d "$default_parameter_info" '.parameterInfo // $d' "$target")"
    api_info="$(jq -r --arg d "$default_api_info" '.apiInfo // $d' "$target")"
  else
    version="0.3.0"
    description=""
    no_code_metadata=""
    parameter_info="$default_parameter_info"
    api_info="$default_api_info"
  fi

  if [[ -z "$description" ]]; then
    description="Runs ${script_name}"
  fi

  jq -n \
    --arg version "$version" \
    --arg body "$body" \
    --arg description "$description" \
    --arg noCodeMetadata "$no_code_metadata" \
    --arg parameterInfo "$parameter_info" \
    --arg apiInfo "$api_info" \
    '{
      version: $version,
      body: $body,
      description: $description,
      noCodeMetadata: $noCodeMetadata,
      parameterInfo: $parameterInfo,
      apiInfo: $apiInfo
    }' > "$target"

  echo "Wrote $target (ResearchLog.runFunction -> \"$script_name\")"
done