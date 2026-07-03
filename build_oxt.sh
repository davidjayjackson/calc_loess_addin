#!/usr/bin/env bash
# Regenerates oxt/LOESS/*.xba from src/*.bas and repacks CalcLoessAddin.oxt.
# Run this after editing anything under src/.
set -euo pipefail

cd "$(dirname "$0")"

wrap_xba() {
    local module_name="$1" src_file="$2" out_file="$3"
    {
        printf '<?xml version="1.0" encoding="UTF-8"?>\n'
        printf '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
        printf '<script:module xmlns:script="http://openoffice.org/2000/script" script:name="%s" script:language="StarBasic" script:moduleType="normal">REM  *****  BASIC  *****\n' "$module_name"
        sed -e 's/&/\&amp;/g' -e 's/</\&lt;/g' -e 's/>/\&gt;/g' "$src_file"
        printf '</script:module>'
    } > "$out_file"
}

mkdir -p oxt/LOESS oxt/META-INF

wrap_xba "Module1" src/LOESS.bas oxt/LOESS/Module1.xba
wrap_xba "Module2" src/SelfTest.bas oxt/LOESS/Module2.xba

# LibreOffice expects a dialog.xlb alongside script.xlb for every Basic
# library, even when the library defines no dialogs - without it, loading
# the library fails with "Error loading BASIC ... dialog.xlb: General Error".
cat > oxt/LOESS/dialog.xlb <<'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="LOESS" library:readonly="false" library:passwordprotected="false"/>
EOF

rm -f CalcLoessAddin.oxt
cd oxt
zip -r -X ../CalcLoessAddin.oxt description.xml META-INF LOESS -x '.*' >/dev/null
cd ..

echo "Built CalcLoessAddin.oxt"
