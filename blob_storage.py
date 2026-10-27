"""Keeps the Excel file in Vercel Blob storage when the app runs on Vercel.

Vercel servers can't keep files on disk, so the workbook is stored as a private blob.
Every save uploads a new copy under a fresh URL and removes the old ones, so a
read never gets an old cached version.
"""
import json
import os
import urllib.parse
import urllib.request

API = "https://vercel.com/api/blob"
API_VERSION = "11"
PREFIX = "cash_desk"


def enabled():
    return bool(os.environ.get("BLOB_READ_WRITE_TOKEN"))


def _request(url, method="GET", data=None, headers=None):
    h = {"authorization": "Bearer " + os.environ["BLOB_READ_WRITE_TOKEN"], "x-api-version": API_VERSION}
    h.update(headers or {})
    req = urllib.request.Request(url, data=data, method=method, headers=h)
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read()


def _list():
    q = urllib.parse.urlencode({"prefix": PREFIX, "limit": 100})
    blobs = json.loads(_request(API + "?" + q))["blobs"]
    return sorted(blobs, key=lambda b: b["uploadedAt"])


_cache = {"url": None, "data": None}


def load_bytes():
    """Latest workbook bytes, or None if nothing is stored yet."""
    blobs = _list()
    if not blobs:
        return None
    url = blobs[-1]["url"]
    if _cache["url"] != url:
        _cache["data"] = _request(url)
        _cache["url"] = url
    return _cache["data"]


def save_bytes(data):
    old = [b["url"] for b in _list()]
    q = urllib.parse.urlencode({"pathname": PREFIX + ".xlsx"})
    res = json.loads(_request(API + "/?" + q, "PUT", data, {
        "x-vercel-blob-access": "private",
        "x-add-random-suffix": "1",
        "x-content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }))
    _cache["url"], _cache["data"] = res["url"], data
    old = [u for u in old if u != res["url"]]
    if old:
        _request(API + "/delete", "POST", json.dumps({"urls": old}).encode(),
                 {"content-type": "application/json"})
