# Rush 1 — Dataviz Studio (TikTok)

**Objectif** : identifier ce qui distingue les contenus et les créateurs TikTok les plus performants, et livrer à un manager non technicien un classeur Excel qui rend ces facteurs évidents en quelques minutes, en version **client**, conforme au RGPD.

**Livrable** : [`deliverable/Dataviz_Studio_TikTok_Client.xlsx`](deliverable/Dataviz_Studio_TikTok_Client.xlsx). Il s'ouvre sur la **Synthèse**, qui donne les conclusions clés et leurs chiffres sans avoir à parcourir les données. Le classeur est protégé en lecture seule (feuilles et structure).

| Onglet | Contenu |
|---|---|
| Synthèse | 5 conclusions chiffrées, leaders par sujet, recommandations, limites |
| Facteurs | Leviers que le créateur contrôle (avec ou sans effet), conséquences du succès |
| Gros comptes | Réponse à l'objection « les gros comptes gagneraient de toute façon » |
| Catégories & TreeMap | Sujets de contenu, leaders par sujet et TreeMap (taille = vues, couleur = sujet) |
| Méthode & limites | Définition du succès, extraits utilisés, méthodes, biais, données personnelles |
| Données | Données préparées et pseudonymisées, source des formules et des tableaux croisés |

**Données** : 300 vidéos issues de 3 des 5 extraits fournis (collecte de 2021).

## Structure

| Dossier | Contenu | Versionné |
|---|---|---|
| `data_raw/` | Les 5 CSV bruts (jamais modifiés) | Non : données personnelles, voir `data_raw/README.md` |
| `scripts/` | Préparation, analyse et vérification des chiffres | Oui |
| `work/` | Fichiers intermédiaires privés (table nettoyée, résultats) | Non |
| `deliverable/` | Uniquement le classeur client | Oui |

## Reproduire l'analyse

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
# placer les 5 CSV du module (rush-1/) dans data_raw/
.venv/bin/python scripts/01_clean.py           # tri et nettoyage des extraits, variables dérivées -> work/
.venv/bin/python scripts/02_analyse.py         # statistiques, gros comptes, leaders -> work/resultats.json
.venv/bin/python scripts/03_verify_numbers.py  # chiffres de la Synthèse = calcul Python (PASS/FAIL)
```

## Données personnelles

Les extraits contiennent des identifiants, pseudos, bios (e-mails, téléphones, comptes de paiement, mentions de santé) et liens concernant de vraies personnes.

- **Agence** : les CSV bruts et les fichiers de travail restent hors du dépôt (`.gitignore`). La table de travail ne conserve ni bios, ni avatars, ni liens, ni noms affichés.
- **Version client** : seuls les 20 créateurs recommandés sont nommés, par leur pseudo public. Les autres sont pseudonymisés (« Créateur 001 »).
- **Justification complète** : onglet *Méthode & limites*, section 6.
