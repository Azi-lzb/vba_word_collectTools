#!/bin/sh
# Run in a disposable Debian/Ubuntu container with /project mounted read-only,
# and an empty writable /out. Requires verification/fixtures from run_all.py.
set -eu
project=/project/rust-cli
tar -xzf "$project/dist/Word汇总工具-0.1.0-linux-x86_64.tar.gz" -C /out
portable=/out/Word汇总工具-0.1.0-linux-x86_64
"$portable/Word汇总工具" --version
run_four() {
    cli=$1
    folder=$2
    mkdir -p "$folder"
    for kind in legacy content-control; do
        if [ "$kind" = legacy ]; then pattern=Rich_Legacy; else pattern=Rich_CC; fi
        for to in excel word; do
            if [ "$to" = excel ]; then ext=xlsx; else ext=docx; fi
            /usr/bin/env PATH=/missing "$cli" extract --backend rust --kind "$kind" --to "$to" \
                --input "$project/verification/fixtures/${pattern}_1.docx" "$project/verification/fixtures/${pattern}_2.docx" \
                --output "$folder/$kind-$to.$ext"
        done
    done
}
run_four "$portable/Word汇总工具" /out/tar
dpkg-deb --info "$project/dist/Word汇总工具_0.1.0_amd64.deb"
dpkg -i "$project/dist/Word汇总工具_0.1.0_amd64.deb"
run_four /usr/bin/Word汇总工具 /out/deb
/usr/bin/env PATH=/missing /usr/bin/Word汇总工具 extract --backend auto --kind content-control --to word \
    --config /usr/share/doc/vbaword-cli/vbaword.example.toml \
    --input "$project/verification/fixtures/Rich_CC_1.docx" --output /out/auto.docx
/usr/bin/env PATH=/missing /usr/bin/Word汇总工具 probe
menu_dir=$(mktemp -d /out/menu-check.XXXXXX)
cd "$menu_dir"
printf '1\n1\n4\n%s\n0\n' "$project/verification/fixtures/Rich_Legacy_1.docx" | /usr/bin/env PATH=/missing /usr/bin/Word汇总工具
test -s 'output/旧式窗体汇总.xlsx'
dpkg -r vbaword-cli
test ! -e /usr/bin/Word汇总工具
echo 'Linux package installation, extraction and removal passed'
