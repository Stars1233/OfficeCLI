#!/bin/bash
set -e

PROJECT="src/officecli/officecli.csproj"
TARGETS="osx-arm64:officecli-mac-arm64 osx-x64:officecli-mac-x64 linux-x64:officecli-linux-x64 win-x64:officecli-win-x64.exe"

build_config() {
    local CONFIG="$1"
    local OUTPUT="bin/$(echo "$CONFIG" | tr '[:upper:]' '[:lower:]')"

    rm -rf "$OUTPUT"
    mkdir -p "$OUTPUT"

    for target in $TARGETS; do
        RID="${target%%:*}"
        NAME="${target##*:}"
        TMPDIR=$(mktemp -d)

        echo "[$CONFIG] Building $RID -> $NAME"
        dotnet publish "$PROJECT" -c "$CONFIG" -r "$RID" -o "$TMPDIR" --nologo -v quiet

        if [ -f "$TMPDIR/officecli.exe" ]; then
            cp "$TMPDIR/officecli.exe" "$OUTPUT/$NAME"
        else
            cp "$TMPDIR/officecli" "$OUTPUT/$NAME"
        fi
        cp "$TMPDIR/officecli.pdb" "$OUTPUT/${NAME%.*}.pdb"

        rm -rf "$TMPDIR"
    done

    rm -rf src/officecli/bin src/officecli/obj

    echo ""
    echo "$CONFIG build complete:"
    ls -lh "$OUTPUT"
}

CONFIG="${1:-release}"

case "$CONFIG" in
    release|Release)
        build_config "Release"
        ;;
    debug|Debug)
        build_config "Debug"
        ;;
    all)
        build_config "Release"
        echo ""
        build_config "Debug"
        ;;
    *)
        echo "Usage: ./build.sh [release|debug|all]"
        echo "  release  - Build Release only"
        echo "  debug    - Build Debug only"
        echo "  all      - Build both (default)"
        exit 1
        ;;
esac
