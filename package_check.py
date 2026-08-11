"""package_check.py -- verify an OPC package's relationships all resolve.

python-pptx will happily open a .pptx whose parts reference targets that
aren't in the package; PowerPoint will not. That asymmetry cost a real deck:
a case study import left a chart pointing at parts that were never written,
every structural assertion passed, and PowerPoint refused the file with a
flat "could not open".

`check_package(path)` returns a list of problems, each a plain sentence. An
empty list means every relationship in every part resolves to something that
exists and every content type is declared.
"""

import posixpath
import re
import zipfile
from xml.etree import ElementTree as ET

_RELS_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_CT_NS = "{http://schemas.openxmlformats.org/package/2006/content-types}"
_RID_ATTR = re.compile(r'r:(?:id|embed|link)="([^"]+)"')


def _rels_path_for(part_name):
    directory, _, filename = part_name.rpartition("/")
    return f"{directory}/_rels/{filename}.rels" if directory else f"_rels/{filename}.rels"


def check_package(path):
    """Every dangling relationship, undeclared part and unresolved rId."""
    problems = []
    with zipfile.ZipFile(path) as package:
        names = set(package.namelist())

        # 1. Every relationship target must exist.
        targets_by_part = {}
        for rels_name in sorted(n for n in names if n.endswith(".rels")):
            owner_dir = posixpath.dirname(posixpath.dirname(rels_name))
            owner = rels_name.replace("/_rels/", "/").removesuffix(".rels")
            try:
                root = ET.fromstring(package.read(rels_name))
            except ET.ParseError as exc:
                problems.append(f"{rels_name} is not parseable XML ({exc}).")
                continue
            mapping = {}
            for rel in root.findall(f"{_RELS_NS}Relationship"):
                rid, target = rel.get("Id"), rel.get("Target", "")
                mapping[rid] = target
                if rel.get("TargetMode") == "External":
                    continue
                resolved = posixpath.normpath(posixpath.join(owner_dir, target))
                if resolved not in names:
                    problems.append(
                        f"{owner or 'package'} relationship {rid} points at "
                        f"{target}, which isn't in the file.")
            targets_by_part[owner] = mapping

        # 2. Every rId referenced inside an XML part must be declared in that
        #    part's own .rels -- the failure mode when ids are remapped but a
        #    relationship isn't carried across.
        for part_name in sorted(n for n in names
                                if n.endswith(".xml") and not n.endswith(".rels")):
            try:
                blob = package.read(part_name).decode("utf-8", "ignore")
            except KeyError:
                continue
            used = set(_RID_ATTR.findall(blob))
            if not used:
                continue
            declared = set(targets_by_part.get(part_name, {}))
            missing = sorted(used - declared)
            if missing:
                problems.append(
                    f"{part_name} refers to {', '.join(missing)} but "
                    f"{'declares ' + ', '.join(sorted(declared)) if declared else 'has no relationships'}.")

        # 3. Every part needs a declared content type.
        if "[Content_Types].xml" in names:
            root = ET.fromstring(package.read("[Content_Types].xml"))
            defaults = {d.get("Extension", "").lower()
                        for d in root.findall(f"{_CT_NS}Default")}
            overrides = {o.get("PartName", "").lstrip("/")
                         for o in root.findall(f"{_CT_NS}Override")}
            for part_name in sorted(names):
                if part_name.endswith(".rels") or part_name == "[Content_Types].xml":
                    continue
                extension = part_name.rsplit(".", 1)[-1].lower()
                if part_name not in overrides and extension not in defaults:
                    problems.append(
                        f"{part_name} has no content type declared.")
        else:
            problems.append("The package has no [Content_Types].xml.")

    return problems


if __name__ == "__main__":
    import sys
    for target in sys.argv[1:]:
        found = check_package(target)
        print(f"\n{target}: {'OK' if not found else str(len(found)) + ' problem(s)'}")
        for problem in found:
            print("  -", problem)
    raise SystemExit(0)
