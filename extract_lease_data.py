import json

path = r"C:\Users\Trey\.claude\projects\C--Users-Trey-OneDrive---Robbie-D--Wood--Inc-Cloaude\caf97d6e-66fc-464e-83c9-68c7f2d6ccf3\tool-results\artifact-d70dd295-1785986373-383f.html"
with open(path, encoding="utf-8") as f:
    content = f.read()

marker = "const LEASE_DATA = "
start = content.find(marker) + len(marker)

i = start
depth = 0
in_str = False
esc = False
while i < len(content):
    c = content[i]
    if in_str:
        if esc:
            esc = False
        elif c == "\\":
            esc = True
        elif c == '"':
            in_str = False
    else:
        if c == '"':
            in_str = True
        elif c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                i += 1
                break
    i += 1

raw = content[start:i]
print("extracted length:", len(raw))
data = json.loads(raw)
print("parsed OK, keys:", list(data.keys()))
print("units count:", len(data["units"]))

out = r"C:\Users\Trey\OneDrive - Robbie D. Wood, Inc\Cloaude\rdw-eos-model\output\lease_data.json"
with open(out, "w", encoding="utf-8") as f:
    f.write(raw)
print("wrote to", out)
