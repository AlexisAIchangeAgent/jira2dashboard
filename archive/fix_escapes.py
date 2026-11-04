"""Replace any non-ASCII character in a VBA file by its JSON escape (backslash + u + 4 hex digits)."""
import sys
from pathlib import Path

path = Path(sys.argv[1])
raw = path.read_bytes().decode("utf-8", errors="replace")
bs = chr(92)
fixed = "".join(c if ord(c) < 128 else bs + "u" + format(0xE9 if c == chr(0xFFFD) else ord(c), "04x") for c in raw)
path.write_bytes(fixed.encode("ascii"))
print("non-ASCII left:", sum(1 for c in fixed if ord(c) > 127))
