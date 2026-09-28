#!/usr/bin/env python3
"""静岡県の若年（15〜29歳）日本人の人口移動データを e-Stat から取得し、可視化用データを生成する。

出典
  - 総務省「住民基本台帳人口移動報告」2025年 年報（実数）
      表11-1 年齢（5歳階級）、男女別他市区町村からの転入者数（市区町村）   statInfId=000040428748
      表11-2 年齢（5歳階級）、男女別他市区町村への転出者数（市区町村）     statInfId=000040428749
      表11-3 年齢（5歳階級）、男女別転入超過数（市区町村）               statInfId=000040428750
      表8-2  年齢（5歳階級）、男女、移動後の住所地別転出者数（都道府県）   statInfId=000040407049
      表8-3  年齢（5歳階級）、男女、住所地別転入超過数（都道府県）         statInfId=000040407050
  - 総務省「住民基本台帳に基づく人口、人口動態及び世帯数調査」
      令和7年1月1日 【日本人住民】市区町村別年齢階級別人口             statInfId=000040306662

使い方
  pip install openpyxl
  python scripts/build_data.py            # data/ 以下に data.js と CSV を出力
"""
from __future__ import annotations

import csv
import json
import os
import sys
import urllib.request
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "scripts" / "raw"
OUT = ROOT / "data"

FILES = {
    "in_muni": "000040428748",
    "out_muni": "000040428749",
    "net_muni": "000040428750",
    "out_dest": "000040407049",
    "net_dest": "000040407050",
    "pop": "000040306662",
}
AGES = ["15歳～19歳", "20歳～24歳", "25歳～29歳"]
AGE_KEYS = ["15", "20", "25"]
SEXES = {"男": "m", "女": "f"}
WARDS = {"22101", "22102", "22103", "22138", "22139", "22140"}  # 静岡市・浜松市の区（市計を使う）

# 本分析での地域区分（伊豆地域は記事の「伊豆地域」に合わせた11市町）
REGION = {
    "伊豆": ["熱海市", "伊東市", "下田市", "伊豆市", "伊豆の国市", "東伊豆町", "河津町", "南伊豆町", "松崎町", "西伊豆町", "函南町"],
    "東部": ["沼津市", "三島市", "富士宮市", "富士市", "御殿場市", "裾野市", "清水町", "長泉町", "小山町"],
    "中部": ["静岡市", "島田市", "焼津市", "藤枝市", "牧之原市", "吉田町", "川根本町"],
    "西部": ["浜松市", "磐田市", "掛川市", "袋井市", "湖西市", "御前崎市", "菊川市", "森町"],
}
REGION_OF = {m: r for r, ms in REGION.items() for m in ms}


def fetch(stat_id: str) -> Path:
    RAW.mkdir(parents=True, exist_ok=True)
    path = RAW / f"{stat_id}.xlsx"
    if not path.exists():
        url = f"https://www.e-stat.go.jp/stat-search/file-download?statInfId={stat_id}&fileKind=0"
        print("download", url, file=sys.stderr)
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        path.write_bytes(urllib.request.urlopen(req, timeout=300).read())
    return path


def rows(stat_id: str):
    ws = openpyxl.load_workbook(fetch(stat_id), read_only=True).active
    return list(ws.iter_rows(values_only=True))


def muni_table(stat_id: str) -> dict:
    """表11-x：{(code, name): {(sex, agekey): value}}（日本人移動者）"""
    r = rows(stat_id)
    sex, age = r[3], r[4]
    idx = {(sex[i], age[i]): i for i in range(7, len(sex))}
    out = {}
    for row in r[6:]:
        code = str(row[5])
        if code.startswith("22") and row[2] == "日本人移動者" and code not in WARDS:
            out[(code, row[6])] = {
                (SEXES[s], k): row[idx[(s, a)]] for s in SEXES for a, k in zip(AGES, AGE_KEYS)
            }
    return out


def dest_table(stat_id: str) -> dict:
    """表8-x：静岡県（日本人移動者）の移動先・移動元別 {pref: {(sex, agekey): value}}"""
    r = rows(stat_id)
    sex, age = r[4], r[5]
    idx = {(sex[i], age[i]): i for i in range(9, len(sex))}
    out = {}
    for row in r[7:]:
        if str(row[1]) == "22000" and row[4] == "日本人移動者":
            name = row[8]
            if name in out:
                continue
            out[name] = {(SEXES[s], k): row[idx[(s, a)]] for s in SEXES for a, k in zip(AGES, AGE_KEYS)}
    return out


def pop_table() -> tuple[dict, dict]:
    """令和7年1月1日 日本人住民の年齢階級別人口：市区町村 {code5: {(sex, agekey): pop}}, 県・全国"""
    r = rows(FILES["pop"])
    hdr = r[1]
    col = {k: hdr.index(a) for a, k in zip(AGES, AGE_KEYS)}
    muni, agg = {}, {}
    sexmap = {"男": "m", "女": "f"}
    for row in r[3:]:
        if row[3] not in sexmap:
            continue
        s = sexmap[row[3]]
        vals = {(s, k): row[c] for k, c in col.items()}
        if row[1] == "合計":
            agg.setdefault("全国", {}).update(vals)
        elif str(row[0]).startswith("22"):
            code = str(row[0])[:5]
            if code == "22000":
                agg.setdefault("静岡県", {}).update(vals)
            else:
                muni.setdefault(code, {}).update(vals)
    return muni, agg


def main() -> None:
    net = muni_table(FILES["net_muni"])
    inn = muni_table(FILES["in_muni"])
    out = muni_table(FILES["out_muni"])
    pop, agg = pop_table()

    munis = []
    for (code, name), v in net.items():
        if code == "22000":
            continue
        rec = {"code": code, "name": name, "region": REGION_OF[name]}
        for s in "mf":
            for k in AGE_KEYS:
                rec[f"{s}{k}_net"] = v[(s, k)]
                rec[f"{s}{k}_in"] = inn[(code, name)][(s, k)]
                rec[f"{s}{k}_out"] = out[(code, name)][(s, k)]
                rec[f"{s}{k}_pop"] = pop[code][(s, k)]
        munis.append(rec)
    assert len(munis) == 35, len(munis)

    dnet = dest_table(FILES["net_dest"])
    dout = dest_table(FILES["out_dest"])
    prefs = [p for p in dnet if p.endswith(("都", "道", "府", "県")) and p != "静岡県"]
    assert len(prefs) == 46
    dests = []
    for p in prefs:
        rec = {"pref": p}
        for s in "mf":
            for k in AGE_KEYS:
                rec[f"{s}{k}_net"] = dnet[p][(s, k)]
                rec[f"{s}{k}_out"] = dout[p][(s, k)]
        dests.append(rec)

    # 検算：35市町の合計 = 県外との収支（県内移動は相殺される）
    for s in "mf":
        for k in AGE_KEYS:
            a = sum(m[f"{s}{k}_net"] for m in munis)
            b = sum(d[f"{s}{k}_net"] for d in dests)
            assert a == b, (s, k, a, b)

    popagg = {
        area: {f"{s}{k}": agg[area][(s, k)] for s in "mf" for k in AGE_KEYS} for area in ("静岡県", "全国")
    }

    data = {
        "meta": {
            "year": 2025,
            "population_date": "2025-01-01",
            "subject": "日本人移動者・15〜29歳",
            "sources": [
                "総務省「住民基本台帳人口移動報告」2025年 年報 表8-2・8-3・11-1〜11-3",
                "総務省「住民基本台帳に基づく人口、人口動態及び世帯数調査」令和7年1月1日（日本人住民）",
            ],
        },
        "municipalities": munis,
        "destinations": dests,
        "population": popagg,
    }
    OUT.mkdir(exist_ok=True)
    (OUT / "data.json").write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    (OUT / "data.js").write_text("window.DATA = " + json.dumps(data, ensure_ascii=False) + ";\n", encoding="utf-8")

    # 人が読むための CSV
    with open(OUT / "municipalities.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        head = ["市町コード", "市町", "地域"]
        for s, sl in (("m", "男"), ("f", "女")):
            for k, kl in zip(AGE_KEYS, ("15-19", "20-24", "25-29")):
                head += [f"{sl}{kl}_転入超過", f"{sl}{kl}_転入", f"{sl}{kl}_転出", f"{sl}{kl}_人口"]
        w.writerow(head)
        for m in sorted(munis, key=lambda x: x["code"]):
            row = [m["code"], m["name"], m["region"]]
            for s in "mf":
                for k in AGE_KEYS:
                    row += [m[f"{s}{k}_net"], m[f"{s}{k}_in"], m[f"{s}{k}_out"], m[f"{s}{k}_pop"]]
            w.writerow(row)
    with open(OUT / "destinations.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        head = ["都道府県"]
        for s, sl in (("m", "男"), ("f", "女")):
            for k, kl in zip(AGE_KEYS, ("15-19", "20-24", "25-29")):
                head += [f"{sl}{kl}_転入超過", f"{sl}{kl}_静岡県からの転出"]
        w.writerow(head)
        for d in dests:
            row = [d["pref"]]
            for s in "mf":
                for k in AGE_KEYS:
                    row += [d[f"{s}{k}_net"], d[f"{s}{k}_out"]]
            w.writerow(row)
    print("ok:", len(munis), "municipalities,", len(dests), "prefectures", file=sys.stderr)


if __name__ == "__main__":
    main()
