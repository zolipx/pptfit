"""Public test decks (Apache POI test data), split into chunks for a CI matrix.

python bench/corpus_list.py N   ->  prints  matrix=[["url", ...], ...]  for $GITHUB_OUTPUT
Decks are downloaded at run time and never stored in this repository.
"""
import json, os, sys, urllib.request

API = "https://api.github.com/repos/apache/poi/contents/test-data/slideshow"
req = urllib.request.Request(API, headers={"Accept": "application/vnd.github+json",
                                           **({"Authorization": "Bearer " + os.environ["GH_TOKEN"]} if os.environ.get("GH_TOKEN") else {})})
items = json.load(urllib.request.urlopen(req))
urls = sorted(i["download_url"] for i in items if i["name"].lower().endswith(".pptx") and 0 < i["size"] < 4_000_000)
n = max(1, int(sys.argv[1]))
chunks = [urls[k::n] for k in range(n) if urls[k::n]]
print("matrix=" + json.dumps(chunks))
print("decks: %d in %d chunks" % (len(urls), len(chunks)), file=sys.stderr)
