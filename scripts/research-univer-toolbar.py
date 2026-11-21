"""Collect public Univer evidence without installing or executing the SDK.

Run with Python 3.12+. Downloads stay in the git-ignored .research.local directory.
The manifest records HTTP failures as well as hashes of successful responses.
"""

import concurrent.futures
import hashlib
import html
import io
import json
from pathlib import Path, PurePosixPath
import re
import subprocess
import tarfile
import urllib.error
import urllib.request


ROOT = Path(__file__).resolve().parents[1]
CACHE = ROOT / ".research.local" / "univer"
VERSION = "1.5.2"
COMMIT = "6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68"
REFERENCE = Path("E:/rust_test/univer")


def fetch(url, destination):
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        body = destination.read_bytes()
        return {"url": url, "status": 200, "sha256": hashlib.sha256(body).hexdigest(),
                "bytes": len(body), "cache": destination.relative_to(ROOT).as_posix()}
    try:
        request = urllib.request.Request(url, headers={"User-Agent": "OfficeViewer-ReferenceAudit/1.0"})
        with urllib.request.urlopen(request, timeout=45) as response:
            body = response.read()
            status = response.status
        destination.write_bytes(body)
        return {"url": url, "status": status, "sha256": hashlib.sha256(body).hexdigest(),
                "bytes": len(body), "cache": destination.relative_to(ROOT).as_posix()}
    except (urllib.error.URLError, TimeoutError) as error:
        return {"url": url, "status": getattr(error, "code", None), "error": str(error)}


def batch(jobs):
    with concurrent.futures.ThreadPoolExecutor(max_workers=6) as executor:
        return list(executor.map(lambda job: fetch(*job), jobs))


def collect():
    CACHE.mkdir(parents=True, exist_ok=True)
    seed_paths = [f"/zh-CN/guides/{product}/features/core" for product in ("sheets", "docs")]
    seed_paths += [f"/zh-CN/overview/features/{product}" for product in ("sheets", "docs")]
    records = batch([(f"https://docs.univer.ai{path}", CACHE / "pages" / (path.replace("/", "_") + ".html"))
                     for path in seed_paths])
    paths = set(seed_paths)
    for record in records:
        if "cache" not in record:
            continue
        body = (ROOT / record["cache"]).read_text(encoding="utf-8")
        paths.update(html.unescape(path) for path in re.findall(
            r'href="(/zh-CN/(?:guides/(?:sheets|docs)/features|overview/features/shared)[^"#?]*)', body))
    records += batch([(f"https://docs.univer.ai{path}", CACHE / "pages" / (path.replace("/", "_") + ".html"))
                      for path in sorted(paths - set(seed_paths))])
    records += batch([(url, CACHE / name) for url, name in [
        ("https://office.univer.ai/?filter=sheets", "office.html"),
        ("https://registry.npmjs.org/-/v1/search?text=scope:univerjs-pro&size=250", "registry-search.json"),
    ]])
    package_names = set()
    for page in (CACHE / "pages").glob("*.html"):
        body = html.unescape(page.read_text(encoding="utf-8"))
        package_names.update(re.findall(r"@univerjs-pro/[a-z][a-z0-9-]+", body))
    search_file = CACHE / "registry-search.json"
    if search_file.exists():
        search = json.loads(search_file.read_text(encoding="utf-8"))
        for entry in search.get("objects", []):
            name = entry["package"]["name"]
            if name.startswith("@univerjs-pro/") and re.search(
                    r"/(?:docs|sheets|drawing|chart|collaboration|edit-history|thread-comment|license|ui|cross-document)", name):
                package_names.add(name)
    manifests = {}
    pending = package_names
    while pending:
        jobs = [(f"https://registry.npmjs.org/{name.replace('/', '%2f')}/{VERSION}",
                 CACHE / "manifests" / (name.split("/")[1] + ".json")) for name in sorted(pending)]
        fetched = batch(jobs)
        records += fetched
        pending = set()
        for record in fetched:
            if "cache" not in record:
                continue
            manifest = json.loads((ROOT / record["cache"]).read_text(encoding="utf-8"))
            manifests[manifest["name"]] = manifest
            for field in ("dependencies", "peerDependencies"):
                pending.update(name for name in manifest.get(field, {})
                               if name.startswith("@univerjs-pro/") and name not in manifests)
        pending -= {record["url"].split("/")[-2].replace("%2f", "/") for record in fetched
                    if "error" in record}
    records += batch([(manifest["dist"]["tarball"], CACHE / "tarballs" / (name.split("/")[1] + ".tgz"))
                      for name, manifest in sorted(manifests.items())])
    extracted = []
    for name, manifest in sorted(manifests.items()):
        tar_path = CACHE / "tarballs" / (name.split("/")[1] + ".tgz")
        if not tar_path.exists():
            continue
        archive = tar_path.read_bytes()
        integrity = manifest["dist"].get("integrity", "")
        import base64
        if integrity.startswith("sha512-"):
            actual = "sha512-" + base64.b64encode(hashlib.sha512(archive).digest()).decode("ascii")
            if actual != integrity:
                raise ValueError(f"Integrity mismatch for {name}")
        with tarfile.open(fileobj=io.BytesIO(archive), mode="r:gz") as package:
            for member in package.getmembers():
                relative = PurePosixPath(member.name)
                if not member.isfile() or relative.is_absolute() or ".." in relative.parts:
                    continue
                product_package = name.startswith(("@univerjs-pro/sheets-", "@univerjs-pro/docs-")) or name in {
                    "@univerjs-pro/chart-ui", "@univerjs-pro/chart-embed", "@univerjs-pro/thread-comment-ui",
                    "@univerjs-pro/edit-history-ui", "@univerjs-pro/print", "@univerjs-pro/watermark",
                }
                is_menu_schema = product_package and "/menu/" in member.name and member.name.endswith(".js")
                if not (member.name.endswith(".d.ts") or is_menu_schema or member.name in (
                        "package/package.json", "package/lib/es/index.js", "package/lib/locale/zh-CN.js")):
                    continue
                target = CACHE / "packages" / name.split("/")[1] / Path(*relative.parts[1:])
                target.parent.mkdir(parents=True, exist_ok=True)
                source = package.extractfile(member)
                assert source is not None
                data = source.read()
                target.write_bytes(data)
        extracted.append({"name": name, "version": manifest["version"], "tarball": manifest["dist"]["tarball"],
                          "integrity": integrity, "sha256": hashlib.sha256(archive).hexdigest(),
                          "license": manifest.get("license"), "bytes": len(archive)})
    # Read the pinned tree directly. The user's reference checkout is left intact.
    files = subprocess.check_output(["git", "-C", str(REFERENCE), "ls-tree", "-r", "--name-only", COMMIT], text=True).splitlines()
    for path in files:
        if not (path.endswith((".ts", ".tsx", ".json")) and path.startswith(("packages/", "presets/", "examples/src/"))):
            continue
        if "/__tests__/" in path or "/__test__/" in path:
            continue
        destination = CACHE / "oss" / path
        if not destination.exists():
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_bytes(subprocess.check_output(["git", "-C", str(REFERENCE), "show", f"{COMMIT}:{path}"]))
    result = {"date": "2026-11-21", "version": VERSION, "pinnedCommit": COMMIT,
              "checkoutCommit": subprocess.check_output(["git", "-C", str(REFERENCE), "rev-parse", "HEAD"], text=True).strip(),
              "resources": records, "packages": extracted}
    (CACHE / "resources.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"pages": len(paths), "packages": len(extracted),
                      "failures": [record for record in records if "error" in record]}, ensure_ascii=False))


if __name__ == "__main__":
    collect()
