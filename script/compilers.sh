#!/usr/bin/env bash
#
# compilers.sh — compile un llms.txt complet (offline, liens déjà inline) en arborescence docs/
#                une feature par fichier + _category_.json Docusaurus par dossier de catégorie.
#
# Usage :
#   script/compilers.sh [llms.txt]   Compile le llms.txt donné (défaut: script/llms.txt) vers docs/ + index llms.txt à la racine.
#   script/compilers.sh --from-docs  Reconstruit le llms.txt complet depuis docs/ existant.
#   script/compilers.sh --reassemble  Reassemble script/llms.txt depuis script/llms.txt.part-* (si le dépôt a été poussé en parties).
#
# En environnement fermé : aucun fetch externe, tout est déjà dans le llms.txt.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
DOCS_DIR="$ROOT_DIR/docs"

slugify() {
  echo "$1" | tr '[:upper:]' '[:lower:]' | sed -e 's/[^a-z0-9 _-]//g' -e 's/ /-/g' -e 's/--*/-/g' -e 's/^-//' -e 's/-$//' | cut -c1-80
}

reassemble() {
  if [ ! -f "$SCRIPT_DIR/llms.txt" ] && compgen -G "$SCRIPT_DIR/llms.txt.part-*" > /dev/null; then
    cat "$SCRIPT_DIR"/llms.txt.part-* > "$SCRIPT_DIR/llms.txt"
    echo "llms.txt réassemblé depuis les parties ($(wc -c < "$SCRIPT_DIR/llms.txt") octets)."
  fi
}

compile_llms() {
  local input="$1"
  [ -f "$input" ] || { echo "llms.txt introuvable: $input" >&2; exit 1; }
  reassemble
  rm -rf "$DOCS_DIR"
  mkdir -p "$DOCS_DIR"
  awk -v docs="$DOCS_DIR" '
    function slug(t) {
      s = tolower(t)
      gsub(/[^a-z0-9 _-]/, "", s)
      gsub(/ /, "-", s)
      gsub(/-+/, "-", s)
      sub(/^-/, "", s); sub(/-$/, "", s)
      return substr(s, 1, 80)
    }
    /^## CAT: / { cat = substr($0, 9); gsub(/^ +| +$/, "", cat); next }
    /^### PAGE: / {
      if (file != "") close(file)
      title = substr($0, 11); gsub(/^ +| +$/, "", title)
      dir = (cat == "_root" || cat == "") ? docs : docs "/" cat
      system("mkdir -p " dir)
      file = dir "/" slug(title) ".md"
      print "---\ntitle: \"" title "\"\n---\n" > file
      print "# " title > file
      next
    }
    /^---$/ { next }
    file != "" { print > file }
    END { if (file != "") close(file) }
  ' "$input"
  # _category_.json dans chaque dossier de catégorie
  (cd "$DOCS_DIR" && find . -type d) | while read -r d; do
    [ "$d" = "." ] && continue
    label="$(basename "$d" | sed 's/-/ /g; s/\b\w/\u&/g')"
    printf '{\n  "label": "%s",\n  "position": 1,\n  "collapsed": false\n}\n' "$label" > "$DOCS_DIR/$d/_category_.json"
  done
  # index llms.txt à la racine
  {
    echo "# Teams SDK - TypeScript Documentation (offline)"
    echo ""
    echo "> Édition offline de https://microsoft.github.io/teams-sdk/llms_docs/llms_typescript.txt (+ contenu complet)."
    echo "> Chaque entrée pointe vers un fichier local de docs/ — aucun fetch externe requis."
    (cd "$DOCS_DIR" && find . -type d) | while read -r d; do
      [ "$d" = "." ] && continue
      cat="${d#./}"
      echo ""
      echo "## $(basename "$cat" | sed 's/-/ /g; s/\b\w/\u&/g')"
      echo ""
      find "$DOCS_DIR/$cat" -maxdepth 1 -name '*.md' | sort | while read -r f; do
        t="$(sed -n 's/^title: "\(.*\)"$/\1/p' "$f" | head -1)"
        echo "- [$t]($cat/$(basename "$f"))"
      done
    done
  } > "$ROOT_DIR/llms.txt"
  echo "Compilation terminée: $(find "$DOCS_DIR" -name '*.md' | wc -l) fichiers dans $DOCS_DIR, index: $ROOT_DIR/llms.txt"
}

from_docs() {
  [ -d "$DOCS_DIR" ] || { echo "docs/ introuvable" >&2; exit 1; }
  out="$SCRIPT_DIR/llms.txt"
  {
    echo "# Teams SDK - TypeScript Documentation (Complete, offline)"
    echo ""
    echo "> Complete offline edition: every documentation page is inlined below as its own section."
    echo "> Internal links are in-file \`#anchors\`; no external fetching is required."
    (cd "$DOCS_DIR" && find . -type d) | while read -r d; do
      [ "$d" = "." ] && continue
      cat="${d#./}"
      echo ""
      echo "## CAT: $cat"
      echo ""
      find "$DOCS_DIR/$cat" -name '*.md' ! -name '_category_.json' | sort | while read -r f; do
        t="$(sed -n 's/^title: "\(.*\)"$/\1/p' "$f" | head -1)"
        echo "### PAGE: $t"
        echo ""
        sed -e '1,4d' "$f"   # retire le frontmatter
        echo ""
        echo "---"
        echo ""
      done
    done
  } > "$out"
  echo "llms.txt complet reconstruit: $out ($(wc -c < "$out") octets)"
}

case "${1:-}" in
  --from-docs)  from_docs ;;
  --reassemble) reassemble ;;
  "")           reassemble; compile_llms "$SCRIPT_DIR/llms.txt" ;;
  -h|--help)    sed -n '2,12p' "${BASH_SOURCE[0]}" ;;
  *)            compile_llms "$1" ;;
esac
