#!/usr/bin/env bash
# Installs LOESS.bas (and SelfTest.bas) into your personal LibreOffice
# "Standard" Basic library (~/.config/libreoffice/4/user/basic/Standard).
#
# Why not just the .oxt extension? LibreOffice's Calc formula engine only
# resolves a bare =FUNCTION(...) name against the user's own "Standard"
# library - a Basic library shipped inside an extension (regardless of what
# it's named) is registered and runnable as a macro, but is never searched
# when compiling a cell formula. This script does the same thing as
# manually pasting the code into Tools > Macros > Edit Macros, just safely
# and repeatably.
#
# Existing content in your Standard library is never removed - this only
# adds two new modules (LOESSAddin, LOESSSelfTest), and backs up the whole
# library first. Re-running this script is safe (it won't duplicate
# entries) and is how you pick up any future updates to LOESS.bas.
set -euo pipefail

cd "$(dirname "$0")"

if pgrep -f soffice.bin >/dev/null 2>&1; then
    echo "LibreOffice appears to be running. Please close it fully first," >&2
    echo "otherwise it will overwrite these changes when it exits." >&2
    exit 1
fi

STANDARD_DIR="${HOME}/.config/libreoffice/4/user/basic/Standard"
mkdir -p "$STANDARD_DIR"

if [ -n "$(ls -A "$STANDARD_DIR" 2>/dev/null)" ]; then
    BACKUP_DIR="${STANDARD_DIR}.backup.$(date +%Y%m%d%H%M%S)"
    cp -r "$STANDARD_DIR" "$BACKUP_DIR"
    echo "Backed up existing Standard library to: $BACKUP_DIR"
fi

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

wrap_xba "LOESSAddin" src/LOESS.bas "$STANDARD_DIR/LOESSAddin.xba"
wrap_xba "LOESSSelfTest" src/SelfTest.bas "$STANDARD_DIR/LOESSSelfTest.xba"

DLB="$STANDARD_DIR/dialog.xlb"
if [ ! -f "$DLB" ]; then
    cat > "$DLB" <<'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" library:readonly="false" library:passwordprotected="false"/>
EOF
fi

XLB="$STANDARD_DIR/script.xlb"
if [ ! -f "$XLB" ]; then
    cat > "$XLB" <<'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" library:readonly="false" library:passwordprotected="false">
</library:library>
EOF
fi

for MOD in LOESSAddin LOESSSelfTest; do
    if ! grep -q "library:name=\"$MOD\"" "$XLB"; then
        sed -i "s#</library:library>#  <library:element library:name=\"$MOD\"/>\n</library:library>#" "$XLB"
        echo "Registered module $MOD"
    else
        echo "Module $MOD already registered (updated code, left script.xlb as-is)"
    fi
done

echo ""
echo "Done. Start LibreOffice and try =LOESS(1;2;1;1;0;0) in a Calc cell (expect 2)."
echo "Run Tools > Macros > Run Macro... > My Macros > Standard > LOESSSelfTest > RunSelfTest to self-check."
