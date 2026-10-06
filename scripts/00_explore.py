"""Phase 1 — Exploration des 5 extraits bruts (lecture seule de data_raw/).

Usage : .venv/bin/python scripts/00_explore.py
Affiche un profil par extrait : forme, types, manquants, doublons, plages
de dates, statistiques des métriques, et le recouvrement entre extraits.
"""
from pathlib import Path

import pandas as pd

RAW = Path(__file__).resolve().parent.parent / "data_raw"
FILES = {
    "trending_videos": "trending_videos.csv",
    "trending_authors": "trending_authors.csv",
    "liked_videos": "tiktok_collected_liked_videos.csv",
    "official_videos": "tiktok_collected_videos.csv",
    "funny_hashtag": "tiktok_funny_hashtag_videos.csv",
}
METRICS = ["n_likes", "n_shares", "n_comments", "n_plays", "video_length",
           "video_stats", "video_shareCount", "video_commentCount",
           "video_playCount", "video_duration", "author_followerCount",
           "author_heartCount", "author_videoCount"]


def load(name):
    return pd.read_csv(RAW / FILES[name], encoding="utf-8")


def profile(name, df):
    print(f"\n{'=' * 70}\n{name}  ({FILES[name]})\n{'=' * 70}")
    print(f"lignes={len(df)}  colonnes={df.shape[1]}")
    print(pd.DataFrame({"dtype": df.dtypes.astype(str),
                        "manquants": df.isna().sum(),
                        "uniques": df.nunique()}).to_string())
    print(f"lignes entièrement dupliquées : {df.duplicated().sum()}")
    for key in ("video_id", "Author Unique ID", "author_uniqueId", "user_name"):
        if key in df.columns:
            print(f"doublons sur {key} : {df[key].duplicated().sum()}")
    for tcol in ("video_time",):
        if tcol in df.columns:
            t = pd.to_datetime(df[tcol], unit="s", utc=True)
            print(f"{tcol} : {t.min():%Y-%m-%d} -> {t.max():%Y-%m-%d}")
    if "video_id" in df.columns:
        # Les identifiants TikTok encodent l'horodatage de publication (bits hauts).
        t = pd.to_datetime(df["video_id"].astype("int64") // 2**32, unit="s", utc=True)
        print(f"date encodée dans video_id : {t.min():%Y-%m-%d} -> {t.max():%Y-%m-%d}")
    for col in ("author_avatarThumb", "Avatar Thumbnail", "video_cover"):
        if col in df.columns:
            exp = df[col].astype(str).str.extract(r"x-expires=(\d+)")[0].dropna()
            if len(exp):
                e = pd.to_datetime(exp.astype("int64"), unit="s", utc=True)
                print(f"expiration URL {col} (≈ date de collecte) : "
                      f"{e.min():%Y-%m-%d %H:%M} -> {e.max():%Y-%m-%d %H:%M}")
    num = [c for c in METRICS if c in df.columns]
    if num:
        print(df[num].describe(percentiles=[.25, .5, .75]).T
              .map(lambda v: f"{v:,.0f}").to_string())
        print("valeurs à 0 :", {c: int((df[c] == 0).sum()) for c in num})


def main():
    dfs = {n: load(n) for n in FILES}
    for n, df in dfs.items():
        profile(n, df)

    print(f"\n{'=' * 70}\nRECOUVREMENTS\n{'=' * 70}")
    vid = {
        "trending_videos": set(dfs["trending_videos"]["video_id"]),
        "liked_videos": set(dfs["liked_videos"]["video_id"]),
        "official_videos": set(dfs["official_videos"]["video_id"]),
        "funny_hashtag": set(dfs["funny_hashtag"]["video_id"]),
    }
    names = list(vid)
    for i, a in enumerate(names):
        for b in names[i + 1:]:
            print(f"video_id {a} ∩ {b} : {len(vid[a] & vid[b])}")

    tv_users = set(dfs["trending_videos"]["user_name"])
    ta_users = set(dfs["trending_authors"]["Author Unique ID"])
    fh_users = set(dfs["funny_hashtag"]["author_uniqueId"])
    lk_users = set(dfs["liked_videos"]["user_name"])
    print(f"auteurs trending_authors uniques : {len(ta_users)}")
    print(f"auteurs trending_videos uniques  : {len(tv_users)}")
    print(f"trending_videos.user_name ∈ trending_authors : "
          f"{len(tv_users & ta_users)}/{len(tv_users)}")
    print(f"vidéos trending appariables à un auteur : "
          f"{dfs['trending_videos']['user_name'].isin(ta_users).sum()}"
          f"/{len(dfs['trending_videos'])}")
    print(f"auteurs funny ∩ trending_videos : {len(fh_users & tv_users)}")
    print(f"auteurs funny ∩ trending_authors : {len(fh_users & ta_users)}")
    print(f"auteurs liked ∩ trending_videos : {len(lk_users & tv_users)}")
    print(f"auteurs funny uniques : {len(fh_users)}")

    fh = dfs["funny_hashtag"]
    print("\nfunny_hashtag : author_heartCount == author_heart ?",
          (fh["author_heartCount"] == fh["author_heart"]).all())
    print("funny_hashtag : video_duration == music_duration ?",
          (fh["video_duration"] == fh["music_duration"]).mean().round(2))
    for c in ("author_verification", "author_privateAccount", "video_isAd",
              "video_originalItem", "video_officialItem", "video_secret",
              "video_forFriend", "video_stitchEnabled", "video_shareEnabled",
              "music_originality", "video_format", "video_quality",
              "video_definition"):
        print(f"  {c}: {fh[c].value_counts(dropna=False).to_dict()}")
    ta = dfs["trending_authors"]
    for c in ("Verified?", "Private Account?"):
        print(f"trending_authors {c}: {ta[c].value_counts().to_dict()}")
    lk = dfs["liked_videos"]
    print("liked_videos : vidéos du compte officiel :",
          (lk["user_name"] == "tiktok").sum())


if __name__ == "__main__":
    main()
