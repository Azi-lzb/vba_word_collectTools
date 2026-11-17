# Release packaging

`package.py` uses Python 3.11+ standard libraries to package prebuilt executables.
It produces a standalone Windows EXE, a Windows ZIP containing only `Word汇总工具.exe` and `config.xlsx`,
a Linux portable tar.gz, a Debian package and SHA256 checksums in `dist/`.
The Python COM worker is embedded; users do not need the source tree.

Local `config.xlsx` files are not committed because they may contain document passwords. If no configuration is available, the Windows packager creates one using the built executable's `init-workbook-config` command. Existing local configuration is preserved.

Windows build:

```powershell
cargo build --release --locked
```

Linux x86_64 build, on a Linux build machine with Rust:

```sh
cargo build --release --locked
cargo test --locked
```

For a broadly compatible glibc binary, build on glibc 2.28 (or use an equivalent
compatible toolchain), and check `readelf --version-info` before shipping.
The current Linux release was built on Rocky Linux 8 / glibc 2.28, then tested
by installing and removing the .deb and running the tar.gz binary in disposable
Ubuntu 20.04 and 22.04 containers. It targets amd64, not ARM64 or Alpine/musl.

Package both prebuilt binaries from Windows:

```powershell
python packaging/package.py --windows-binary target/release/vbaword-cli.exe --linux-binary target/linux-x86_64/release/vbaword-cli
```

Either binary option may be supplied on its own. Linux archive executables have
mode 0755. The Debian package installs `/usr/bin/Word汇总工具` and documentation
under `/usr/share/doc/vbaword-cli`. It does not install Python, Office or modify
VBA/workbooks. COM is a Windows-only optional backend.

Install on Debian/Ubuntu:

```sh
sudo apt install ./Word汇总工具_0.1.0_amd64.deb
Word汇总工具 --help
```

Use the portable Linux package:

```sh
tar -xzf Word汇总工具-0.1.0-linux-x86_64.tar.gz
cd Word汇总工具-0.1.0-linux-x86_64
./Word汇总工具 extract --kind content-control --to word --backend rust --input /path/to/forms --output /path/to/summary.docx
```

`verify-linux.sh` is the container smoke test (requires generated verification
fixtures). Mount the project at `/project` read-only and an empty output folder
at `/out`, then run `sh /project/rust-cli/packaging/verify-linux.sh`. Packaging
validation results are retained in `dist/package-verification.json`.

Release filenames and portable executables use Word汇总工具. The Debian package identifier remains `vbaword-cli` for package-manager compatibility.
