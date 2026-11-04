#!/usr/bin/env python3
"""Inspect the public structure of one or more Google Forms (read-only, never submits).

Reads form URLs from environment variables GFORM_URL_1..GFORM_URL_9, fetches each
viewform page anonymously, parses the embedded FB_PUBLIC_LOAD_DATA_ array and prints
the question titles, types, entry ids and options. Form ids and URLs are never printed.
"""
import base64
import hashlib
import json
import os
import re
import sys
import urllib.request
import urllib.error

TYPE_NAMES = {
    0: "short_answer", 1: "paragraph", 2: "radio", 3: "dropdown", 4: "checkbox",
    5: "linear_scale", 6: "section_text", 7: "grid", 8: "page_break", 9: "date",
    10: "time", 11: "image", 12: "video", 13: "file_upload",
}

UA = ("Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 "
      "(KHTML, like Gecko) Version/18.0 Mobile/15E148 Safari/604.1")


def fetch(url: str):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept-Language": "zh-TW,zh;q=0.9"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.status, resp.geturl(), resp.read().decode("utf-8", "replace")


def classify_final_url(final_url: str) -> str:
    if "accounts.google.com" in final_url:
        return "SIGN_IN_REQUIRED"
    if final_url.endswith("closedform") or "/closedform" in final_url:
        return "CLOSED"
    if "viewform" in final_url:
        return "OPEN_VIEWFORM"
    return "OTHER"


def seal(text: str, key: str) -> str:
    """Obfuscate text with a keyed keystream so the log never shows it in clear.

    Only the holder of the (masked) key input can read it back. Not meant as
    strong cryptography; it keeps form ids out of a public log.
    """
    if not key:
        return "(no key given)"
    data = text.encode("utf-8")
    stream = b""
    counter = 0
    while len(stream) < len(data):
        stream += hashlib.sha256(f"{key}:{counter}".encode("utf-8")).digest()
        counter += 1
    out = bytes(a ^ b for a, b in zip(data, stream))
    return base64.urlsafe_b64encode(out).decode("ascii")


def extract_load_data(html: str):
    m = re.search(r"FB_PUBLIC_LOAD_DATA_\s*=\s*(\[.+?\])\s*;\s*</script>", html, re.S)
    if not m:
        m = re.search(r"FB_PUBLIC_LOAD_DATA_\s*=\s*(\[.+\])\s*;", html, re.S)
    if not m:
        return None
    return json.loads(m.group(1))


def describe(data):
    out = {}
    form = data[1] if len(data) > 1 and isinstance(data[1], list) else None
    if form is None:
        return {"error": "unexpected layout: data[1] missing", "top_len": len(data)}
    out["title"] = form[8] if len(form) > 8 else None
    out["description"] = form[0] if len(form) > 0 else None
    out["top_len"] = len(data)
    out["sign_in_flag_data18"] = data[18] if len(data) > 18 else None
    out["settings_block_form10"] = form[10] if len(form) > 10 else None
    questions = []
    for el in (form[1] or []):
        try:
            q = {
                "item_id": el[0],
                "title": el[1],
                "type": TYPE_NAMES.get(el[3], f"type_{el[3]}"),
                "type_code": el[3],
            }
            entries = el[4] if len(el) > 4 and el[4] else []
            q["entries"] = []
            for e in entries:
                ent = {"entry_id": e[0], "required": bool(e[2]) if len(e) > 2 else None}
                opts = e[1] if len(e) > 1 and e[1] else []
                labels = []
                has_other = False
                for o in opts:
                    label = o[0] if isinstance(o, list) and o else o
                    if isinstance(o, list) and len(o) > 4 and o[4]:
                        has_other = True
                    labels.append(label)
                ent["options"] = labels
                ent["has_other_option"] = has_other
                if len(e) > 7 and e[7]:
                    ent["date_time_flags"] = e[7]
                q["entries"].append(ent)
            questions.append(q)
        except Exception as ex:  # keep going, show raw shape
            questions.append({"error": str(ex), "raw_prefix": json.dumps(el, ensure_ascii=False)[:300]})
    out["questions"] = questions
    return out


def main():
    results = {}
    key = os.environ.get("GFORM_SEAL_KEY", "")
    for i in range(1, 10):
        url = os.environ.get(f"GFORM_URL_{i}", "").strip()
        label = os.environ.get(f"GFORM_LABEL_{i}", f"form{i}").strip() or f"form{i}"
        if not url:
            continue
        rec = {"label": label}
        try:
            status, final_url, html = fetch(url)
            rec["http_status"] = status
            rec["final_url_kind"] = classify_final_url(final_url)
            rec["final_path_is_viewform"] = final_url.rstrip("/").split("?")[0].endswith("viewform")
            rec["final_url_sealed"] = seal(final_url.split("?")[0], key)
            rec["has_fbzx_input"] = bool(re.search(r'name="fbzx"', html))
            rec["has_pageHistory_input"] = bool(re.search(r'name="pageHistory"', html))
            rec["has_emailAddress_input"] = bool(re.search(r'name="emailAddress"', html))
            rec["mentions_record_email"] = ("記錄你的電子郵件" in html) or ("記錄您的電子郵件" in html) or ("will be recorded" in html)
            data = extract_load_data(html)
            if data is None:
                rec["error"] = "FB_PUBLIC_LOAD_DATA_ not found"
                rec["html_len"] = len(html)
                rec["title_tag"] = (re.search(r"<title>(.*?)</title>", html, re.S) or [None, None])[1]
            else:
                rec["form"] = describe(data)
        except urllib.error.HTTPError as e:
            rec["error"] = f"HTTPError {e.code}"
            rec["error_url_sealed"] = seal((e.geturl() or "").split("?")[0], key)
        except Exception as e:
            rec["error"] = f"{type(e).__name__}: {e}"
        results[label] = rec
    if not key:
        print("GFORM_SEAL_KEY is required; refusing to print form structure in clear", file=sys.stderr)
        return 1
    print("=== GFORM_INSPECT_BEGIN (sealed) ===")
    print(seal(json.dumps(results, ensure_ascii=False), key))
    print("=== GFORM_INSPECT_END ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
