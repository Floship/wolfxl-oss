# Floship WolfXL releases

This fork builds WolfXL Community wheels for Portal. The wheel and native module share a version with a `+floship.N` suffix. `Cargo.lock` pins Rust dependencies.

- `load_workbook(..., read_only=True, disk_shared_strings=True)` keeps the shared-string index on disk. It supports paths and binary streams.
- `Workbook(write_only=True, inline_strings=True)` writes text inline, so the writer can release each row without retaining a workbook-wide string table.

Build from the exact candidate commit with `bash scripts/build-floship-wheel.sh <absolute-output-directory> <rust-target> <cpython-3.14.9-path>`. The script pins Rust, Maturin and the Linux cross-compiler. Install the target Rust standard libraries with `rustup target add --toolchain 1.96.0 <rust-target>`. Build the Linux wheels on a Linux host, and the macOS wheel on Apple Silicon.

Install each candidate wheel in a clean CPython 3.14 environment. Run the bounded streaming tests and community wheel smoke check against the installed package. Verify real workbook and large-file results before publishing. Publish the qualified wheels and their SHA-256 checksums on an immutable GitHub release at the exact source commit. Pin the release assets and hashes in Portal's lockfile.
