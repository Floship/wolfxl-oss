#!/usr/bin/env bash
set -euo pipefail

if [[ $# != 3 ]]; then
    echo "Usage: $0 <output-directory> <rust-target> <cpython-3.14.9-path>" >&2
    exit 2
fi

output=$1
target=$2
python=$3
cd "$(dirname "$0")/.."

export RUSTUP_TOOLCHAIN=1.96.0
export CARGO_PROFILE_RELEASE_LTO=thin
export CARGO_PROFILE_RELEASE_CODEGEN_UNITS=8
export CARGO_PROFILE_RELEASE_STRIP=none

"$python" -c 'import platform; assert platform.python_version() == "3.14.9"; assert platform.python_implementation() == "CPython"'

args=(--locked --release --target "$target" --interpreter "$python" --out "$output")
case "$target" in
    x86_64-unknown-linux-gnu|aarch64-unknown-linux-gnu)
        uvx --with ziglang==0.15.2 --from maturin==1.14.1 maturin build "${args[@]}" --zig --compatibility manylinux_2_28
        ;;
    aarch64-apple-darwin)
        export MACOSX_DEPLOYMENT_TARGET=11.0
        uvx --from maturin==1.14.1 maturin build "${args[@]}"
        ;;
    *)
        echo "Unsupported wheel target: $target" >&2
        exit 2
        ;;
esac
