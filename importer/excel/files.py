"""File-level helpers shared by the Excel readers."""
import hashlib
import os
import re

EXPORT_PREFIX = re.compile(r"^[0-9a-f]{12}_")  # the upload folder prefixes copies with a content hash


def unique_contents(paths, skipped):
    """One path per distinct file content (copies kept under another name are read once)."""
    seen, keep = {}, []
    for p in sorted(os.fspath(p) for p in paths):
        try:
            with open(p, "rb") as f:
                digest = hashlib.sha256(f.read()).hexdigest()
        except OSError as exc:
            skipped.append({"file": p, "error": f"unreadable: {exc}"})
            continue
        if digest in seen:
            skipped.append({"file": p, "error": f"identical to {os.path.basename(seen[digest])}"})
        else:
            seen[digest] = p
            keep.append(p)
    return keep


def file_display_name(path) -> str:
    """The file name without the 12-hex-digit prefix the folder export puts in front of some."""
    return EXPORT_PREFIX.sub("", os.path.basename(path))
