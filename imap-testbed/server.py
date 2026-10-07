#!/usr/bin/env python3
"""A tiny read-only IMAPS server that serves ./sample as INBOX, for testing Margince's IMAP connector.

  python3 server.py --init     make cert.pem/key.pem and a throwaway credentials.txt (once)
  python3 server.py            serve on 127.0.0.1:1993

Implements only what the connector sends: CAPABILITY, LOGIN, LIST, SELECT/EXAMINE, FETCH, UID FETCH,
NOOP, LOGOUT. Messages are the sample's NNNN.eml files; the number in the name is the UID.
"""
import argparse
import asyncio
import glob
import os
import re
import secrets
import ssl
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
UIDVALIDITY = 1
TOKEN = re.compile(r'"((?:[^"\\]|\\.)*)"|(\S+)')


def init_files():
    cert, key, cred = (os.path.join(HERE, n) for n in ("cert.pem", "key.pem", "credentials.txt"))
    subprocess.run(
        ["openssl", "req", "-x509", "-newkey", "ec", "-pkeyopt", "ec_paramgen_curve:prime256v1", "-nodes",
         "-keyout", key, "-out", cert, "-days", "30", "-subj", "/CN=localhost",
         "-addext", "subjectAltName=DNS:localhost,IP:127.0.0.1", "-addext", "basicConstraints=critical,CA:TRUE"],
        check=True, capture_output=True)
    os.chmod(key, 0o600)
    with open(os.path.join(HERE, "sample", "manifest.json"), encoding="utf-8") as f:
        import json
        owner = json.load(f)["owner"]
    fd = os.open(cred, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w") as f:
        f.write(f"host=localhost\nport=1993\nuser={owner}\npassword={secrets.token_urlsafe(18)}\n")
    print(f"wrote {cert}, {key}, {cred}")


def load_credentials():
    out = {}
    with open(os.path.join(HERE, "credentials.txt"), encoding="utf-8") as f:
        for line in f:
            k, _, v = line.strip().partition("=")
            out[k] = v
    return out


def load_messages():
    msgs = {}
    for path in sorted(glob.glob(os.path.join(HERE, "sample", "*.eml"))):
        with open(path, "rb") as f:
            msgs[int(os.path.basename(path)[:-4])] = f.read()
    return msgs


def parse_set(spec, universe):
    """Resolve an IMAP number set ('1:5,9,12:*') against a sorted list; '*' is the last element."""
    if not universe:
        return []
    top = universe[-1]
    want = set()
    for part in spec.split(","):
        a, _, b = part.partition(":")
        lo = top if a == "*" else int(a)
        hi = lo if not b else (top if b == "*" else int(b))
        lo, hi = min(lo, hi), max(lo, hi)
        want.update(n for n in universe if lo <= n <= hi)
    return sorted(want)


def tokens(line):
    return [m.group(1) if m.group(1) is not None else m.group(2) for m in TOKEN.finditer(line)]


class Session:
    def __init__(self, reader, writer, msgs, creds):
        self.r, self.w, self.msgs, self.creds = reader, writer, msgs, creds
        self.uids = sorted(msgs)
        self.authed = False
        self.selected = False

    def send(self, data):
        self.w.write(data if isinstance(data, bytes) else data.encode())

    async def run(self):
        self.send("* OK [CAPABILITY IMAP4rev1] testbed ready\r\n")
        while True:
            raw = await self.r.readline()
            if not raw:
                return
            parts = tokens(raw.decode("utf-8", "replace").rstrip("\r\n"))
            if len(parts) < 2:
                continue
            tag, cmd, args = parts[0], parts[1].upper(), parts[2:]
            if cmd == "UID" and args:
                cmd, args = "UID " + args[0].upper(), args[1:]
            if await self.dispatch(tag, cmd, args, raw.decode("utf-8", "replace")):
                return
            await self.w.drain()

    async def dispatch(self, tag, cmd, args, line):
        if cmd == "CAPABILITY":
            self.send("* CAPABILITY IMAP4rev1\r\n")
        elif cmd == "NOOP":
            pass
        elif cmd == "LOGOUT":
            self.send(f"* BYE bye\r\n{tag} OK LOGOUT completed\r\n")
            await self.w.drain()
            return True
        elif cmd == "LOGIN":
            ok = len(args) >= 2 and args[0] == self.creds["user"] and secrets.compare_digest(args[1], self.creds["password"])
            if not ok:
                self.send(f"{tag} NO [AUTHENTICATIONFAILED] rejected\r\n")
                return False
            self.authed = True
            self.send(f"{tag} OK [CAPABILITY IMAP4rev1] LOGIN completed\r\n")
            return False
        elif not self.authed:
            self.send(f"{tag} BAD login first\r\n")
            return False
        elif cmd == "LIST":
            ref_pat = args[:2]
            if len(ref_pat) == 2 and ref_pat[1] in ("INBOX", "*", "%"):
                self.send('* LIST (\\HasNoChildren) "/" "INBOX"\r\n')
        elif cmd in ("SELECT", "EXAMINE"):
            self.selected = bool(args) and args[0].upper() == "INBOX"
            if not self.selected:
                self.send(f"{tag} NO no such mailbox\r\n")
                return False
            nxt = (self.uids[-1] + 1) if self.uids else 1
            self.send(f"* {len(self.uids)} EXISTS\r\n* 0 RECENT\r\n* FLAGS (\\Seen)\r\n"
                      f"* OK [PERMANENTFLAGS ()]\r\n* OK [UIDVALIDITY {UIDVALIDITY}]\r\n* OK [UIDNEXT {nxt}]\r\n"
                      f"{tag} OK [READ-ONLY] {cmd} completed\r\n")
            return False
        elif cmd in ("FETCH", "UID FETCH"):
            if not self.selected or len(args) < 2:
                self.send(f"{tag} BAD select first\r\n")
                return False
            by_uid = cmd == "UID FETCH"
            universe = self.uids if by_uid else list(range(1, len(self.uids) + 1))
            want_body = "BODY[" in line.upper() or "BODY.PEEK[" in line.upper()
            for n in parse_set(args[0], universe):
                uid = n if by_uid else self.uids[n - 1]
                seq = self.uids.index(uid) + 1
                if want_body:
                    body = self.msgs[uid]
                    self.send(f"* {seq} FETCH (UID {uid} BODY[] {{{len(body)}}}\r\n".encode() + body + b")\r\n")
                else:
                    self.send(f"* {seq} FETCH (UID {uid})\r\n")
        else:
            self.send(f"{tag} BAD unsupported {cmd}\r\n")
            return False
        self.send(f"{tag} OK {cmd} completed\r\n")
        return False


async def main(port):
    msgs, creds = load_messages(), load_credentials()
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ctx.minimum_version = ssl.TLSVersion.TLSv1_2
    ctx.load_cert_chain(os.path.join(HERE, "cert.pem"), os.path.join(HERE, "key.pem"))

    async def handle(reader, writer):
        try:
            await Session(reader, writer, msgs, creds).run()
        except (ConnectionError, asyncio.IncompleteReadError):
            pass
        finally:
            writer.close()

    server = await asyncio.start_server(handle, "127.0.0.1", port, ssl=ctx)
    print(f"serving {len(msgs)} messages on 127.0.0.1:{port}")
    async with server:
        await server.serve_forever()


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--init", action="store_true")
    ap.add_argument("--port", type=int, default=1993)
    a = ap.parse_args()
    if a.init:
        init_files()
    else:
        asyncio.run(main(a.port))
