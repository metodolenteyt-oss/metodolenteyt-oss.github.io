#!/usr/bin/env python3
"""Junta apps-script/*.src.js ao Code.gs e escapa todo caractere nao-ASCII
(\\uXXXX nas strings; letras sem acento nos comentarios) -> apps-script/Code.gs
pode ser colado no Apps Script sem risco de acentos corrompidos."""
import re, os, unicodedata
D = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "apps-script")
base = open(os.path.join(D, "Code.gs"), encoding="utf-8").read()
base = base.split("// ===================================================================== CRM")[0].rstrip() + "\n\n"
src = base + open(os.path.join(D, "crm.src.js"), encoding="utf-8").read()
out = []
for line in src.split("\n"):
    m = re.search(r"(^\s*(/\*\*|\*|//))|(\s//\s)", line)
    cut = m.start() if m else len(line)
    code, com = line[:cut], line[cut:]
    code = "".join(c if ord(c) < 128 else "".join("\\u%04X" % int.from_bytes(c.encode("utf-16-be")[i:i + 2], "big")
                   for i in range(0, len(c.encode("utf-16-be")), 2)) for c in code)
    com = unicodedata.normalize("NFKD", com).encode("ascii", "ignore").decode()
    out.append(code + com)
open(os.path.join(D, "Code.gs"), "w", encoding="ascii").write("\n".join(out))
print("Code.gs:", len(out), "linhas, ASCII")
