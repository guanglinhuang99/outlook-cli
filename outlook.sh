#!/bin/sh
exec node "$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)/bin/outlook.mjs" "$@"
