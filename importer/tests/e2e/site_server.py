"""A throw-away copy of the site for browser tests: the real app, an in-memory CRM instead of a login to a real
one, and a company master of one invented customer (or the real customer lists with --real-master).
Nothing leaves this process. GET /__test/crm tells a test what the in-memory CRM holds.

    .venv/bin/python -m tests.e2e.site_server --port 5099
"""
import argparse
import functools

from flask import jsonify

from excel import run
from excel.companies import Master
from tests.fake_site_crm import SiteCRM
from webui import selftest, server


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, required=True)
    ap.add_argument("--real-master", action="store_true")
    args = ap.parse_args(argv)

    crm = SiteCRM()
    server._client = lambda: crm
    if not args.real_master:
        master = Master.from_records(selftest.MASTER)
        run._master = functools.lru_cache(None)(lambda importer_dir: master)

    @server.app.get("/__test/crm")
    def crm_state():
        return jsonify({"companies": len(crm.companies), "contacts": len(crm.contacts), "leads": len(crm.leads),
                        "activities": len(crm.activities), "writes": len(crm.writes),
                        "fields": sum(len(v) for v in crm.fields.values())})

    server.app.run(host="127.0.0.1", port=args.port, debug=False, threaded=True)


if __name__ == "__main__":
    main()
