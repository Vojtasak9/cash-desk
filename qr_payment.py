"""QR payment codes in the Czech "QR Platba" format (SPAYD).

Every Czech (and most Slovak) banking app can scan these. The code is
generated locally with the free `qrcode` library - no online service,
no account, no fees.  Spec: https://qr-platba.cz/pro-vyvojare/specifikace-formatu/
"""
import io
import re

import qrcode
import qrcode.image.svg


def _iban_checksum_ok(iban):
    s = iban[4:] + iban[:4]
    digits = "".join(str(int(c, 36)) for c in s)
    return int(digits) % 97 == 1


def to_iban(account):
    """Accept an IBAN or a Czech account number like 19-123456789/0800."""
    a = re.sub(r"\s+", "", account or "").upper()
    if not a:
        raise ValueError("Set your bank account in Settings first.")
    if re.fullmatch(r"[A-Z]{2}\d{2}[A-Z0-9]{10,30}", a):
        if not _iban_checksum_ok(a):
            raise ValueError("The IBAN is not valid (checksum).")
        return a
    m = re.fullmatch(r"(?:(\d{1,6})-)?(\d{2,10})/(\d{4})", a)
    if not m:
        raise ValueError("Account must be an IBAN or a Czech account number like 123456789/0800.")
    prefix, number, bank = m.group(1) or "0", m.group(2), m.group(3)
    bban = bank + prefix.zfill(6) + number.zfill(10)
    check = 98 - int(bban + "123500") % 97  # "CZ00" -> C=12, Z=35, 00
    return "CZ%02d%s" % (check, bban)


def _clean(text, limit):
    # '*' separates fields in SPAYD, so it must not appear in values
    return (text or "").replace("*", " ").strip()[:limit]


def spayd(account, amount, currency="CZK", message="", variable_symbol=None, recipient=""):
    parts = ["SPD", "1.0", "ACC:" + to_iban(account), "AM:%.2f" % amount, "CC:" + (currency or "CZK")]
    if message:
        parts.append("MSG:" + _clean(message, 60))
    if recipient:
        parts.append("RN:" + _clean(recipient, 35))
    if variable_symbol:
        parts.append("X-VS:" + str(variable_symbol)[:10])
    return "*".join(parts)


def qr_svg(data):
    img = qrcode.make(data, image_factory=qrcode.image.svg.SvgPathImage, box_size=12, border=2)
    buf = io.BytesIO()
    img.save(buf)
    return buf.getvalue().decode("utf-8")
