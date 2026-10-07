#!/usr/bin/env bash
# One-shot setup for testing Margince's own IMAP capture with a sample of the Thunderbird export.
# Run from anywhere:  bash imap-testbed/setup.sh
# Steps stop at the first failure. Safe to re-run; it resets only the disposable dev database.
set -euo pipefail

TESTBED="$(cd "$(dirname "$0")" && pwd)"
MARGINCE="$TESTBED/../margince"

echo "== 1/6 unit test for the dev-only IMAP switch"
(cd "$MARGINCE/backend" && gofmt -l internal/modules/capture/imap/ && go vet ./internal/modules/capture/imap/ && go test ./internal/modules/capture/imap/)

echo "== 2/6 sample the export (read-only) -> $TESTBED/sample"
python3 "$TESTBED/sample.py"

echo "== 3/6 test server certificate and throwaway login"
python3 "$TESTBED/server.py" --init

echo "== 4/6 tell the dev stack to trust the test server (dev posture only)"
ENVFILE="$MARGINCE/.env.local"
LINE="MARGINCE_DEV_IMAP_TESTBED_CA=\"$TESTBED/cert.pem\""
grep -v '^MARGINCE_DEV_IMAP_TESTBED_CA=' "$ENVFILE" > "$ENVFILE.tmp" || true
echo "$LINE" >> "$ENVFILE.tmp"
chmod 600 "$ENVFILE.tmp"
mv "$ENVFILE.tmp" "$ENVFILE"
grep -c '^ANTHROPIC_API_KEY=.' "$ENVFILE" | sed 's/^1$/   Anthropic key found: yes/;s/^0$/   Anthropic key found: NO (AI stays fake)/'

echo "== 5/6 start the test mail server in the background on 127.0.0.1:1993"
pkill -f "imap-testbed/server.py" 2>/dev/null || true
nohup python3 -u "$TESTBED/server.py" > "$TESTBED/server.log" 2>&1 &
sleep 2
cat "$TESTBED/server.log"

echo "== 6/6 rebuild + restart the dev stack on a fresh, empty database (takes a few minutes)"
(cd "$MARGINCE" && make dev-stop && make dev-fresh && make seed-dev)

echo
echo "Done. Connection values for Settings -> Integrations -> IMAP mailbox:"
grep -v '^password=' "$TESTBED/credentials.txt"
echo "(password is in $TESTBED/credentials.txt)"
