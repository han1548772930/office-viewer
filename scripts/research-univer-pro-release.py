"""Collect published Pro release files for static review without loading SDK code."""

import argparse
import base64
import concurrent.futures
import hashlib
import io
import json
from pathlib import Path, PurePosixPath
import re
import tarfile
import urllib.error
import urllib.request


ROOT = Path(__file__).resolve().parents[1]


def download(url, destination):
    if destination.exists():
        return destination.read_bytes()
    request = urllib.request.Request(url, headers={"User-Agent": "OfficeViewer-StaticReference/1.0"})
    with urllib.request.urlopen(request, timeout=45) as response:
        data = response.read()
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes(data)
    return data


def relevant(name):
    return bool(re.match(r"^@univerjs-pro/(?:docs-|sheets-|chart-|collaboration-(?:client-ui|embed)|"
                         r"edit-history-ui$|embed(?:-ui|-unit-ui)?$|shape-editor-ui$|print$)", name))


def collect_package(name, version, cache):
    slug = name.split("/")[1]
    manifest_url = f"https://registry.npmjs.org/{name.replace('/', '%2f')}/{version}"
    try:
        manifest_data = download(manifest_url, cache / "manifests" / f"{slug}.json")
        manifest = json.loads(manifest_data)
        if manifest.get("name") != name or manifest.get("version") != version:
            raise ValueError(f"Unexpected package identity for {name}")
        data = download(manifest["dist"]["tarball"], cache / "tarballs" / f"{slug}.tgz")
        integrity = manifest["dist"]["integrity"]
        algorithm, expected = integrity.split("-", 1)
        actual = base64.b64encode(hashlib.new(algorithm, data).digest()).decode("ascii")
        if actual != expected:
            raise ValueError(f"Integrity mismatch for {name}")
        extracted = []
        with tarfile.open(fileobj=io.BytesIO(data), mode="r:gz") as archive:
            for member in archive.getmembers():
                path = PurePosixPath(member.name)
                if not member.isfile() or path.is_absolute() or ".." in path.parts or path.parts[0] != "package":
                    continue
                if not (member.name.endswith(".d.ts") or member.name == "package/package.json"
                        or member.name == "package/lib/es/index.js"
                        or ("/menu/" in member.name and member.name.endswith(".js"))):
                    continue
                target = cache / "packages" / slug / Path(*path.parts[1:])
                target.parent.mkdir(parents=True, exist_ok=True)
                source = archive.extractfile(member)
                if source is None:
                    raise ValueError(f"Missing archive file {member.name}")
                target.write_bytes(source.read())
                extracted.append(path.relative_to("package").as_posix())
        return {"name": name, "version": version, "manifest": manifest_url,
                "tarball": manifest["dist"]["tarball"], "integrity": integrity,
                "sha256": hashlib.sha256(data).hexdigest(), "bytes": len(data),
                "files": extracted, "license": manifest.get("license")}
    except (urllib.error.URLError, TimeoutError, ValueError) as error:
        return {"name": name, "version": version, "manifest": manifest_url,
                "error": str(error), "status": getattr(error, "code", None)}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("version")
    options = parser.parse_args()
    if not re.fullmatch(r"\d+\.\d+\.\d+", options.version):
        parser.error("version must be an exact release")
    cache = ROOT / ".research.local" / f"univer-{options.version}"
    cache.mkdir(parents=True, exist_ok=True)
    historical = json.loads((ROOT / ".research.local/univer/resources.json").read_text(encoding="utf-8"))
    names = {item["name"] for item in historical["packages"] if relevant(item["name"])}
    search_url = "https://registry.npmjs.org/-/v1/search?text=scope:univerjs-pro&size=250"
    search = json.loads(download(search_url, cache / "registry-search.json"))
    names.update(item["package"]["name"] for item in search.get("objects", [])
                 if relevant(item["package"]["name"]))
    with concurrent.futures.ThreadPoolExecutor(max_workers=6) as executor:
        results = list(executor.map(lambda name: collect_package(name, options.version, cache), sorted(names)))
    report = {"version": options.version, "discovery": ["historical package names", search_url],
              "packages": results, "sdkExecuted": False,
              "scope": "Sheets, Docs, chart, embed, edit-history, print and shared UI packages"}
    (cache / "release-packages.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"packages": len(results), "successful": sum("error" not in item for item in results),
                      "failures": [item for item in results if "error" in item]}, ensure_ascii=False))


if __name__ == "__main__":
    main()
