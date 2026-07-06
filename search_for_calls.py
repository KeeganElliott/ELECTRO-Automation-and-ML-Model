import win32com.client
from pathlib import Path

ies = win32com.client.Dispatch("IES.Document")

methods = sorted([name for name in dir(ies) if not name.startswith("_")])

out = Path("electro_api_methods_full.txt")
out.write_text("\n".join(methods), encoding="utf-8")

print(f"Saved {len(methods)} methods to: {out.resolve()}")