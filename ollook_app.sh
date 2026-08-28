#!/usr/bin/env bash
# Lance l'utilitaire autonome Ollook.
cd "$(dirname "${BASH_SOURCE[0]}")"
exec python3 ollook_app.py "$@"
