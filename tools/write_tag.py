"""Write a Green Lady barrel link to an NTAG21x tag on the ACS ACR1255U-J1 reader.

  python write_tag.py               -> show what is on the tag now (no changes)
  python write_tag.py B1        -> write https://mrdanxxx.github.io/green-lady/#B1

Tags are never locked: only user-memory pages are written; lock bytes untouched.
"""
import sys
from smartcard.System import readers

BASE = "https://mrdanxxx.github.io/green-lady/#"
PREFIXES = {0: "", 1: "http://www.", 2: "https://www.", 3: "http://", 4: "https://"}


def reader():
    acs = [r for r in readers() if str(r).startswith("ACS")]
    if not acs:
        sys.exit("No ACS reader found. Plug in the ACR1255U-J1 by USB.")
    c = acs[0].createConnection()
    try:
        c.connect()
    except Exception:
        sys.exit("No tag on the reader. Put the tag flat on the reader and try again.")
    return c


def apdu(c, cmd):
    data, sw1, sw2 = c.transmit(cmd)
    if (sw1, sw2) != (0x90, 0x00):
        raise RuntimeError("Reader error %02X %02X for %s" % (sw1, sw2, bytes(cmd[:5]).hex()))
    return data


def read_page(c, p):
    return apdu(c, [0xFF, 0xB0, 0x00, p, 0x04])[:4]


def describe(c):
    uid = apdu(c, [0xFF, 0xCA, 0x00, 0x00, 0x00])
    cc = read_page(c, 3)
    size = cc[2] * 8
    raw = []
    for p in range(4, 4 + size // 4):
        raw += read_page(c, p)
    print("Tag UID:", ":".join("%02X" % b for b in uid), "| user memory", size, "bytes")
    text = None
    if raw and raw[0] == 0x03:
        ln = raw[1]
        rec = raw[2:2 + ln]
        if len(rec) > 4 and rec[3] == ord("U"):
            pl = rec[2]
            body = rec[4:4 + pl]
            text = PREFIXES.get(body[0], "") + bytes(body[1:]).decode("utf-8", "replace")
        elif ln:
            text = "(non-URL data, %d bytes)" % ln
    print("Currently on tag:", text or "(empty)")
    return size, text


def write(c, size, url):
    assert url.startswith("https://")
    body = url[len("https://"):].encode()
    payload = bytes([0x04]) + body
    rec = bytes([0xD1, 0x01, len(payload)]) + b"U" + payload
    tlv = bytes([0x03, len(rec)]) + rec + b"\xFE"
    if len(tlv) > size:
        sys.exit("URL too long for this tag (%d > %d bytes)" % (len(tlv), size))
    tlv += b"\x00" * (-len(tlv) % 4)
    for i in range(0, len(tlv), 4):
        apdu(c, [0xFF, 0xD6, 0x00, 4 + i // 4, 0x04] + list(tlv[i:i + 4]))


if __name__ == "__main__":
    c = reader()
    size, current = describe(c)
    if len(sys.argv) > 1:
        url = BASE + sys.argv[1].upper()
        if current and current != url and "--overwrite" not in sys.argv:
            sys.exit("Tag already has data. Re-run with --overwrite to replace it.")
        write(c, size, url)
        _, check = describe(c)
        print("WRITTEN OK" if check == url else "WRITE CHECK FAILED")
