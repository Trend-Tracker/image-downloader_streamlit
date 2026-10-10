"""
Images to ZIP - Streamlit app.

Run locally:
    pip install streamlit
    streamlit run app.py

Images are fetched by this app on your own computer, straight from their
websites. No browser CORS limits, no proxies, no third-party services.
"""
import io
import mimetypes
import re
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from urllib.parse import unquote, urlparse
import urllib.request

import streamlit as st

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0 Safari/537.36"}

st.set_page_config(page_title="Images to ZIP", page_icon="🗜️", layout="centered")


def parse_links(text: str) -> list[str]:
    seen, links = set(), []
    for token in re.split(r"[\s,]+", text):
        if re.match(r"^https?://\S+$", token, re.I) and token not in seen:
            seen.add(token)
            links.append(token)
    return links


def make_name(url: str, index: int, content_type: str) -> str:
    base = unquote(Path(urlparse(url).path).name) or "image"
    base = re.sub(r"[^\w.\-]+", "_", base)
    if not re.search(r"\.[A-Za-z0-9]{2,5}$", base):
        base += mimetypes.guess_extension((content_type or "").split(";")[0]) or ".jpg"
    return f"{index + 1:03d}_{base}"


def fetch(item):
    index, url = item
    try:
        req = urllib.request.Request(url, headers=HEADERS)
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = resp.read()
            ctype = resp.headers.get("Content-Type", "")
        if ctype.startswith(("text/", "application/json")):
            raise ValueError(f"got a web page instead of an image ({ctype})")
        return {"index": index, "url": url, "name": make_name(url, index, ctype), "data": data, "error": None}
    except Exception as err:
        return {"index": index, "url": url, "name": None, "data": None, "error": str(err)}


st.title("Images to ZIP")
st.write("Paste image links, one per line. Get them all in a single ZIP file.")

raw = st.text_area(
    "Image URLs",
    height=220,
    placeholder="https://example.com/photo1.jpg\nhttps://example.com/photo2.png",
)
links = parse_links(raw)

col1, col2 = st.columns([3, 1], vertical_alignment="bottom")
zip_name = col1.text_input("ZIP file name", value="images")
col2.caption(f"{len(links)} link{'s' if len(links) != 1 else ''}")

if st.button("Prepare ZIP", type="primary", disabled=not links):
    st.session_state.pop("zip_result", None)
    progress = st.progress(0.0, text=f"Fetching 0 of {len(links)}...")
    results = []
    with ThreadPoolExecutor(max_workers=4) as pool:
        for done, res in enumerate(pool.map(fetch, enumerate(links)), start=1):
            results.append(res)
            progress.progress(done / len(links), text=f"Fetching {done} of {len(links)}...")
    progress.empty()

    good = [r for r in results if not r["error"]]
    if good:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as zf:
            for r in good:
                zf.writestr(r["name"], r["data"])
        st.session_state["zip_result"] = {
            "bytes": buf.getvalue(),
            "name": re.sub(r"[^\w\- ]+", "_", zip_name.strip() or "images") + ".zip",
            "report": [
                {"Link": r["url"], "Status": "Added" if not r["error"] else "Failed", "Details": r["name"] or r["error"]}
                for r in results
            ],
            "good": len(good),
            "total": len(results),
        }
    else:
        st.session_state["zip_result"] = {
            "bytes": None,
            "report": [{"Link": r["url"], "Status": "Failed", "Details": r["error"]} for r in results],
            "good": 0,
            "total": len(results),
        }

result = st.session_state.get("zip_result")
if result:
    if result["bytes"]:
        msg = f"{result['good']} of {result['total']} images are ready."
        (st.success if result["good"] == result["total"] else st.warning)(msg)
        st.download_button(
            "Download ZIP",
            data=result["bytes"],
            file_name=result["name"],
            mime="application/zip",
            type="primary",
        )
    else:
        st.error("No images could be downloaded. Check that the links point directly to image files.")
    st.dataframe(result["report"], width="stretch", hide_index=True)