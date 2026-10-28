#!/bin/bash
# 把 pptfit 源码打包传到 Colab 的 /content/pptfit（不在本机运行任何代码），并核对两边一致
set -e
cd "$(dirname "$0")"
zipf="$(mktemp -d)/pptfit_src.zip"
zip -qrX "$zipf" pptfit tests bench README.md -x "*__pycache__*"
want=$(shasum -a 256 "$zipf" | cut -d' ' -f1)
for i in 1 2 3; do
  colab upload -s lab "$zipf" /content/pptfit_src.zip >/dev/null 2>&1 || true
  got=$(echo 'import hashlib,zipfile,shutil,os; d=open("/content/pptfit_src.zip","rb").read(); h=hashlib.sha256(d).hexdigest(); shutil.rmtree("/content/pptfit",ignore_errors=True); zipfile.ZipFile("/content/pptfit_src.zip").extractall("/content/pptfit"); print("SHA", h)' | colab exec -s lab 2>/dev/null | grep -o 'SHA [0-9a-f]*' | cut -d' ' -f2)
  [ "$got" = "$want" ] && { echo "synced ${want:0:12}"; exit 0; }
  sleep 5
done
echo "SYNC FAILED" >&2; exit 1
