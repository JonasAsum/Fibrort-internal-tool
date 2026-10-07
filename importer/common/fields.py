"""Find-or-create a custom field without trusting a label that another importer may have used."""


def _fits(field: dict, type_: str, options: list[str] | None) -> bool:
    if field.get("type") != type_:
        return False
    return not options or set(options) <= set(field.get("options") or [])


def ensure_field(client, object_: str, label: str, type_: str, options: list[str] | None, source: str) -> dict:
    """Reuse the active field called `label` only if it takes our type and every option we will
    write; otherwise use (or create) "<label> (<source>)". A same-named field from another
    importer (Excel "Customer type" = new/existing/dealer) is left alone instead of 422-ing."""
    existing = {f["label"].lower(): f for f in client.custom_fields(object_)}
    for candidate in (label, f"{label} ({source})"):
        found = existing.get(candidate.lower())
        if found and _fits(found, type_, options):
            return found
    taken = label.lower() in existing
    body = {"object": object_, "label": f"{label} ({source})" if taken else label, "type": type_}
    if options:
        body["options"] = options
    return client.create_custom_field(body)
