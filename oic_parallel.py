"""Compatibility CLI for the original round-robin collector.

Catalog/member inventory, paging, curl_cffi, workers and exports now share
the existing all-dimension engine. The original standalone implementation
is retained in Git history (6c17a7e).
"""
from __future__ import annotations

import sys
from oic_all_dim import main

if __name__=="__main__":
    # The same options and saved plan are used by the UI and both CLI names.
    aliases={"--prepare":"--job-plan","--launch":None,"--merge":"--status"}
    sys.argv=[sys.argv[0]]+[aliases.get(arg,arg) for arg in sys.argv[1:]
                           if aliases.get(arg,arg) is not None]
    raise SystemExit(main())
