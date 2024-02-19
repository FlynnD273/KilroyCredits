import pyperclip
import openpyxl
from openpyxl.cell.cell import Cell
import sys
import os

if len(sys.argv) == 1:
    print(f"Please start program with a file input")
    exit()
path = sys.argv[1]
if os.path.exists(path):
    xl = openpyxl.load_workbook(path)
else:
    print(f"Path {path} does not exist!")
    exit()

class skit:
    title: str
    cast: list[tuple[str, str]]

    def __init__(self, title: str):
        self.title = title
        self.cast = []

    def __str__(self) -> str:
        s = f"\t\t{self.title}\n\n"
        s += "\n".join([f"{char}\t{name}" for name, char in self.cast])
        return s

# Hardcoding this because it should be the same between master caster sheets
casting = xl["Casting"]
skits: list[skit] = []

# Accumulated rows
acc: list[tuple[Cell, ...]] = []

for row in casting.iter_rows():
    # Empty row
    if all(map(lambda c: c.value is None, row[:3])):
        # If rows have been accumulated, dump them into a Skit
        if len(acc) > 0:
            # Do stuff here
            sk: skit = skit(str(acc[0][0].value))
            if len(acc) > 1:
                auth = acc[1][2].value
                if auth is not None:
                    sk.cast.append((str(auth).strip(), "Author"))
            if len(acc) > 2:
                prod = acc[2][2].value
                if prod is not None:
                    sk.cast.append((str(prod).strip(), "Producer"))
            if len(acc) > 3:
                ad = acc[3][2].value
                if ad is not None:
                    sk.cast.append((str(ad).strip(), "AD"))

            for line in acc[1:]:
                if line[0].value is None:
                    continue

                sk.cast.append((str(line[0].value).strip(), str(line[1].value).strip()))

            skits.append(sk)
        acc = []
        continue

    acc.append(row)

content = ""
for sk in skits:
    if sk.title == "Template Sketch":
        continue
    content += str(sk)
    content += "\n\n"

print(content)
pyperclip.copy(content)
