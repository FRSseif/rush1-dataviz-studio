"""
 Nettoyage et préparation (version agence, privée).
"""
import re
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data_raw"
WORK = ROOT / "work"

# Date de collecte estimée (dernière publication observée / expiration des URL signées).
COLLECTE = {
    "Tendances": pd.Timestamp("2021-08-25", tz="UTC"),
    "Humour": pd.Timestamp("2021-08-25", tz="UTC"),
    "Likées par @tiktok": pd.Timestamp("2021-08-22", tz="UTC"),
}

# Hashtags « génériques » qui visent l'algorithme plutôt que le sujet de la vidéo.
GENERIQUES = {
    "fyp", "fypシ", "foryou", "foryoupage", "foryourpage", "fy", "4u", "fypage",
    "viral", "viralvideo", "viraltiktok", "trending", "trend", "xyzbca", "zycbca",
    "parati", "pourtoi", "fürdich", "foru", "fyeeeeeeeeeepppppp", "viiral",
    "viral_video", "tiktok", "foryoupagе",
}

# Thème du contenu : mots-clés génériques cherchés dans la légende et le pseudo.
# Ordre = priorité (le sujet de la vidéo prime sur le registre humoristique).
# Sans mot-clé (légende vide, emojis seuls, simple mention) : « Sujet non déclaré ».
THEMES = [
    ("Animaux", r"dog|pupp|\bpets?\b|petsof|petlife|cutepet|\bcat\b|cats|catlover|"
                r"cattiktok|kitt|meow|animal|bird|cockatoo|goffin|raccoon|coon\b|"
                r"\blion\b|duck|frenchie|collie|corgi|llama|кот|hamster|chipmunk|pups?\b|"
                r"dogtok|\bwee\b|scottish|lemur|🐶|🐱|🦝"),
    ("Famille & enfants", r"\bkids?\b|kidjokes|baby|babies|\bson\b|daughter|\bmoms?\b|"
                          r"mommy|momma|mama|\bdad|daddy|parent|family|toddler|"
                          r"brother|sister|grandma|grandparent|pregnant|childhood|"
                          r"momsof|singlemom|mexicanmom|mexicandad|친남매|친오빠|boymom|garner"),
    ("Lifestyle, musique & sport", r"danc|music|\bsongs?\b|beatbox|\bsing(ing|er)?\b|"
                                   r"\bptd\b|permissiontodance|trendingsong|training|"
                                   r"messi|lebron|football|soccer|\bnba\b|basket|"
                                   r"sports?\b|tennis|\bgym\b|neymar|food|recipe|receta|"
                                   r"cook|tostadas|makeup|eyeliner|skincare|\bspa\b|"
                                   r"asmr|fashion|hoodie|outfit|pimple|snack|fruit|"
                                   r"completemylook|photo|magic|drawing"),
    ("Humour & sketchs", r"funny|funnyy|comedy|humou?r|umor|prank|meme|joke|\blol\b|"
                         r"lmao|lmfao|satire|troll|\bfail|\bpov\b|relatable|skit|"
                         r"😂|🤣|😆"),
]

CTA = r"\?|comment|tag |\btag\b|follow|share|reply to|who else|which|dites|qui "
COLLAB = r"@|#duet|#dueto|#stitch|#이어|reply to"


def load_video_extract(fname, source):
    """Extraits au format commun (trending, liked)."""
    df = pd.read_csv(RAW / fname, encoding="utf-8")
    return pd.DataFrame({
        "echantillon": source,
        "createur": df["user_name"],
        "video_id": df["video_id"].astype("int64"),
        "legende": df["video_desc"].fillna(""),
        "publication": pd.to_datetime(df["video_time"], unit="s", utc=True),
        "duree_s": df["video_length"],
        "vues": df["n_plays"],
        "likes": df["n_likes"],
        "commentaires": df["n_comments"],
        "partages": df["n_shares"],
    })


def load_funny():
    df = pd.read_csv(RAW / "tiktok_funny_hashtag_videos.csv", encoding="utf-8")
    out = pd.DataFrame({
        "echantillon": "Humour",
        "createur": df["author_uniqueId"],
        "video_id": df["video_id"].astype("int64"),
        "legende": df["video_desc"].fillna(""),
        # L'horodatage de publication est encodé dans les 32 bits de poids fort de l'ID.
        "publication": pd.to_datetime(df["video_id"].astype("int64") // 2**32,
                                      unit="s", utc=True),
        "duree_s": df["video_duration"],
        "vues": df["video_playCount"],
        "likes": df["video_stats"],  # vérifié : video_stats = nombre de likes
        "commentaires": df["video_commentCount"],
        "partages": df["video_shareCount"],
        "abonnes": df["author_followerCount"],
        "likes_cumules_compte": df["author_heartCount"],
        "nb_videos_compte": df["author_videoCount"],
        "verifie": df["author_verification"],
        "son_original": df["music_originality"],
        "stitch_autorise": df["video_stitchEnabled"],
    })
    # Pseudo du compte : sert seulement à repérer un thème (ex. « …cat… »).
    out["_indice_theme"] = df["author_nickname"].fillna("")
    return out


def theme(text):
    t = text.lower()
    for name, pattern in THEMES:
        if re.search(pattern, t):
            return name
    return "Sujet non déclaré"


def theme_dominant(themes):
    counts = themes.value_counts()
    best = set(counts[counts == counts.max()].index)
    ordre = [name for name, _ in THEMES] + ["Sujet non déclaré"]
    return next(t for t in ordre if t in best)


def classe_duree(s):
    if s <= 10:
        return "1. ≤ 10 s"
    if s <= 15:
        return "2. 11-15 s"
    if s <= 30:
        return "3. 16-30 s"
    if s <= 60:
        return "4. 31-60 s"
    return "5. > 60 s"


JOURS = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]


def creneau(h):
    return ("0-5 h" if h < 6 else "6-11 h" if h < 12 else "12-17 h" if h < 18
            else "18-23 h")


def tranche_abonnes(n):
    if pd.isna(n):
        return np.nan
    if n < 1_000_000:
        return "1. < 1 M"
    if n < 3_000_000:
        return "2. 1-3 M"
    if n < 8_000_000:
        return "3. 3-8 M"
    return "4. ≥ 8 M"


def derive(df):
    hashtags = df["legende"].str.findall(r"#([^\s#]+)")
    df["nb_hashtags"] = hashtags.str.len()
    df["nb_hashtags_generiques"] = hashtags.apply(
        lambda tags: sum(t.lower().strip(".,!") in GENERIQUES for t in tags))
    df["hashtag_generique"] = df["nb_hashtags_generiques"] > 0
    texte = (df["legende"].str.replace(r"#[^\s#]+", " ", regex=True)
             .str.replace(r"@[\w.]+", " ", regex=True)
             .str.replace(r"\s+", " ", regex=True).str.strip())
    df["longueur_texte"] = texte.str.len()
    df["legende_vide"] = df["legende"].str.strip() == ""
    df["format_collab"] = df["legende"].str.contains(COLLAB, case=False, regex=True)
    df["question_ou_appel"] = df["legende"].str.contains(CTA, case=False, regex=True)

    df["jour"] = df["publication"].dt.dayofweek  # 0 = lundi (UTC)
    df["jour_nom"] = df["jour"].map(dict(enumerate(JOURS)))
    df["heure_utc"] = df["publication"].dt.hour
    df["creneau_utc"] = df["heure_utc"].apply(creneau)
    df["week_end"] = df["jour"] >= 5
    ref = df["echantillon"].map(COLLECTE)
    df["age_jours"] = ((ref - df["publication"]).dt.total_seconds() // 86400).astype(int)

    df["classe_duree"] = df["duree_s"].apply(classe_duree)
    df["engagement"] = (df["likes"] + df["commentaires"] + df["partages"]) / df["vues"]
    df["taux_likes"] = df["likes"] / df["vues"]
    df["taux_commentaires"] = df["commentaires"] / df["vues"]
    df["taux_partages"] = df["partages"] / df["vues"]
    if "abonnes" in df:
        df["vues_par_abonne"] = df["vues"] / df["abonnes"]
        df["tranche_abonnes"] = df["abonnes"].apply(tranche_abonnes)

    indice = df["legende"] + " " + df["createur"] + " " + df.get(
        "_indice_theme", pd.Series("", index=df.index)).fillna("")
    df["theme_video"] = indice.apply(theme)
    # Thème du créateur = sujet le plus fréquent de ses vidéos (tous échantillons) ;
    # un sujet déclaré l'emporte sur « Sujet non déclaré » en cas d'égalité.
    df["theme"] = df.groupby("createur")["theme_video"].transform(theme_dominant)

    # Succès = quartile supérieur du taux d'engagement, calculé dans chaque échantillon.
    seuil = df.groupby("echantillon")["engagement"].transform(lambda s: s.quantile(0.75))
    df["top_performer"] = df["engagement"] >= seuil
    return df.drop(columns=[c for c in df.columns if c.startswith("_")])


def main():
    WORK.mkdir(exist_ok=True)
    tv = load_video_extract("trending_videos.csv", "Tendances")
    lk = load_video_extract("tiktok_collected_liked_videos.csv", "Likées par @tiktok")
    fh = load_funny()
    df = derive(pd.concat([tv, fh, lk], ignore_index=True))

    # Contrôles de qualité documentés.
    assert not df.duplicated(["echantillon", "video_id"]).any()
    assert (df["vues"] > 0).all() and df[["likes", "partages", "vues"]].gt(0).all().all()
    print(df.groupby("echantillon").agg(
        videos=("video_id", "size"), createurs=("createur", "nunique"),
        commentaires_nuls=("commentaires", lambda s: int((s == 0).sum())),
        legendes_vides=("legende_vide", "sum")).to_string())
    print(df["theme"].groupby(df["echantillon"]).value_counts().unstack(0).fillna(0)
          .astype(int).to_string())
    df.to_csv(WORK / "videos_clean.csv", index=False, encoding="utf-8")
    print(f"-> {WORK / 'videos_clean.csv'} ({len(df)} lignes)")


if __name__ == "__main__":
    main()
