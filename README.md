# Rush 1 — Dataviz Studio (TikTok)

**Objectif** : identifier ce qui distingue les contenus et les créateurs TikTok les plus performants, et livrer à un manager non technicien un classeur Excel qui rend ces facteurs évidents en quelques minutes, en version **client**, conforme au RGPD.

**Livrable** : [`deliverable/Dataviz_Studio_TikTok_Client.xlsx`](deliverable/Dataviz_Studio_TikTok_Client.xlsx). Il s'ouvre sur la **Synthèse**, qui donne les conclusions clés et leurs chiffres sans avoir à parcourir les données.

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
| `scripts/` | Chaîne reproductible (voir ci-dessous) | Oui |
| `work/` | Fichiers intermédiaires privés (table nettoyée, résultats, contrôles) | Non |
| `deliverable/` | Uniquement le classeur client | Oui |

## Reproduire

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
# placer les 5 CSV du module (rush-1/) dans data_raw/
.venv/bin/python scripts/01_clean.py            # nettoyage, variables dérivées -> work/
.venv/bin/python scripts/02_analyse.py          # tests, gros comptes, leaders -> work/resultats.json
.venv/bin/python scripts/03_build_workbook.py   # classeur client -> deliverable/
.venv/bin/python scripts/04_audit_client_xlsx.py        # audit RGPD et conformité (PASS/FAIL)
.venv/bin/python scripts/05_verify_numbers.py --excel   # chiffres Python = Excel (--excel : Excel pour Mac)
```

`ooxml_extras.py` ajoute les tableaux croisés dynamiques natifs (qu'aucune bibliothèque Python ne sait écrire). `excel_mac.py` fait ouvrir, recalculer et réenregistrer le classeur par Excel pour le valider.

## Étape manuelle : insérer le TreeMap natif

AppleScript ne sait pas créer de TreeMap dans Excel pour Mac. Ses données, ses couleurs (thème du classeur) et son emplacement sont déjà prêts :

1. Ouvrir `deliverable/Dataviz_Studio_TikTok_Client.xlsx` dans Excel, onglet **Catégories & TreeMap**.
2. Sélectionner **B42:D67** (en-têtes *Sujet*, *Tuile*, *Vues cumulées* compris).
3. Ruban **Insertion** > **Insérer un graphique hiérarchique** > **Compartimentage** (Treemap).
4. Déplacer et redimensionner le graphique pour couvrir **F42:K67**.
5. Titre : `Qui capte les vues ? Vues cumulées par créateur (Tendances + Humour), couleur = sujet`.
6. Optionnel : double-cliquer sur une tuile > *Options de la série* > *Étiquettes parent* : **Bannière**.
7. Clic droit > **Modifier le texte de remplacement** :
   `TreeMap : chaque tuile est un créateur ; taille = vues cumulées (Tendances + Humour), couleur = sujet. Les créateurs non retenus sont regroupés par sujet.`
8. Revenir sur l'onglet **Synthèse**, cliquer sur **A1**, enregistrer (⌘S) au format .xlsx, puis fermer Excel.
9. Lancer `.venv/bin/python scripts/06_finaliser.py`. Le script retire ce qu'Excel ajoute à l'enregistrement (nom de l'auteur, chemin local, étiquette de confidentialité, imprimante), remet la Synthèse en premier, puis relance l'audit et la vérification. Il doit afficher **28/28** et **19/19 PASS**.

## Données personnelles

Les extraits contiennent des identifiants, pseudos, bios (e-mails, téléphones, comptes de paiement, mentions de santé) et liens concernant de vraies personnes.

- **Agence** : les CSV bruts et les fichiers de travail restent hors du dépôt (`.gitignore`). La table de travail ne conserve ni bios, ni avatars, ni liens, ni noms affichés.
- **Version client** : seuls les 20 créateurs recommandés sont nommés, par leur pseudo public. Les autres sont pseudonymisés (« Créateur 001 »).
- **Justification complète** : onglet *Méthode & limites*, section 6.
