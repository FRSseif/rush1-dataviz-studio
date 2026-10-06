# Walkthrough — 5 minutes sur le classeur, sans slides

Support : `deliverable/Dataviz_Studio_TikTok_Client.xlsx`. Avant de commencer, ouvrir le fichier dans Excel. Il s'ouvre sur **Synthèse**, en A1.

## Déroulé minuté

### 0:00 – 0:45 · La question, la définition du succès, les données (onglet Synthèse, A1)

- « Le client veut savoir ce qui fait performer un contenu TikTok avant de choisir ses créateurs. Le succès n'étant défini nulle part, je l'ai défini. »
- **Succès = taux d'engagement** = (likes + commentaires + partages) ÷ vues. Trois raisons :
  1. c'est la réaction du public qu'achète une marque ;
  2. c'est un ratio, donc il ne favorise pas mécaniquement les gros comptes ni les vidéos anciennes ;
  3. dans les tendances, il va avec la portée (ρ = +0,21).
- **Top performers** = le quart supérieur de chaque échantillon. Les vues sont toujours affichées à côté.
- **Données** : 3 extraits sur 5, soit 300 vidéos.
  - Les **tendances** sont l'échantillon principal.
  - Le **top 100 #funny** est le seul extrait avec le nombre d'abonnés.
  - Les **vidéos likées par @tiktok** servent de contrôle.
- **Écartés** : les *auteurs tendance*, qui n'ont aucune métrique et ne contiennent que des données personnelles, et les vidéos du *compte officiel*, qui n'est pas un créateur recrutable. Détail dans *Méthode & limites*, section 2.

### 0:45 – 2:45 · Trois conclusions chiffrées (Synthèse D6:D10, puis les onglets de preuve)

1. **La taille du compte ne fait pas la portée** (Synthèse D6 : ρ = +0,28). Cliquer sur « → Gros comptes ». Les comptes de plus de 8 M d'abonnés en ont **36 fois** plus que ceux de moins de 1 M, mais n'ont que **1,3 fois** plus de vues médianes (65 M contre 48 M). Montrer le tableau par tranche (lignes 11-15) et le graphique.
2. **Les « recettes » de publication n'y changent rien** (Synthèse D8 : −0,7 pt pour #fyp/#foryou, p = 0,32). Onglet *Facteurs*, tableau A : la colonne Verdict est grise presque partout (hashtags, jour, heure, appel à l'action, mentions, son original). « Un facteur sans effet est un résultat : c'est le message pour le brief. »
3. **Seul levier qui ressort : des vidéos un peu plus longues** (Synthèse D9 : +5,5 pts).
   - Dans les tendances : 15,6 % d'engagement au-delà de 15 s contre 10,1 % en dessous, et 47 % de top performers contre 20 % (p = 0,02).
   - Mais le verdict est **« Effet isolé »** : rien de tel dans les deux autres échantillons. C'est une piste à tester en A/B, pas une règle.
   - Montrer le graphique *Facteurs* K24 et les effectifs (section C).

### 2:45 – 3:30 · Catégories, leaders et TreeMap (onglet Catégories & TreeMap)

- **5 sujets** déduits des hashtags et légendes (tableau A). Aucun ne gagne partout : 4,2 points d'écart, non significatif (p = 0,33).
- **Leaders** (tableau B) : meilleur taux d'engagement rapporté à la médiane de son échantillon (« indice »), parmi les créateurs qui ont au moins une portée médiane. Exemple : @www9781 en Animaux, à 19,6 %, soit 1,8 fois la médiane. Un leader de portée par sujet (tableau C).
- **TreeMap** (zone F42) : **taille = vues cumulées**, **couleur = sujet**. Ce sont deux informations différentes. Message : 10 % des créateurs captent 42 % des vues.

### 3:30 – 4:15 · L'objection « les gros comptes gagneraient de toute façon » (onglet Gros comptes)

- **Ce qu'on peut dire** :
  - parmi les vidéos qui ont percé, les abonnés n'expliquent qu'une petite partie de la portée (ρ = 0,28) ;
  - l'engagement baisse avec la taille du compte (19,1 % sous 1 M, 13,2 % au-delà de 8 M) ;
  - par abonné, les petits comptes rapportent 25 fois plus de vues.
- **Ce qu'on ne peut pas dire** (lignes 26-29) :
  - la probabilité de percer : on ne voit pas les petits comptes qui ont échoué (biais du survivant) ;
  - la causalité : les abonnés sont mesurés après le succès ;
  - la réponse pour les tendances : pas d'abonnés dans l'extrait ;
  - à vues égales, l'écart d'engagement s'affaiblit (−0,15, non significatif).

### 4:15 – 5:00 · Données personnelles (onglet Méthode & limites, section 6)

- **Base légale** : intérêt légitime de l'agence (art. 6.1.f). Les contenus ont été rendus publics par leurs auteurs, l'usage est limité à l'étude, sans profilage.
- **Minimisation** : la version client ne contient ni identifiants, ni liens, ni photos, ni noms affichés, ni bios (e-mails, téléphones, paiements, santé), ni légendes (qui citent des tiers), ni musiques, ni badges.
- **Pseudonymisation** : seuls les **20 créateurs recommandés** sont nommés, par leur pseudo public. Les autres sont des codes « Créateur 001 », sans table de correspondance dans le fichier. Un pseudo qui révèle un état de santé n'est jamais nommé (art. 9).
- **Durée et droits** : suppression à la fin de la campagne (6 mois au plus). Information des créateurs au premier contact (art. 14) et droit d'opposition (art. 21).
- **Preuve** : le script `04_audit_client_xlsx.py` cherche dans tout le fichier (feuilles, caches de TCD, graphiques, métadonnées) chacune des valeurs exclues des CSV bruts.

## 15 questions probables et réponses courtes

1. **Pourquoi l'engagement et pas les vues ?**
   - Les vues dépendent de la durée de mise en ligne (ρ âge–vues = +0,24 dans les tendances) et de la taille du compte.
   - L'engagement mesure la réaction et dépend peu de l'âge de la vidéo (ρ = −0,15, non significatif).
   - Les vues restent affichées partout.
2. **Pourquoi le quart supérieur comme seuil ?**
   - Avec 100 vidéos, un top 10 % ne ferait que 10 vidéos, trop peu pour comparer.
   - Le quartile donne 25 vidéos par échantillon.
   - Les seuils sont calculés par échantillon (15,6 % dans les tendances, 20,5 % dans l'humour), parce que les échantillons ne sont pas sélectionnés de la même façon.
3. **Pourquoi avoir écarté `trending_authors` ?**
   - Il n'a aucune métrique de performance.
   - Son seul apport serait le badge vérifié, qui est une conséquence du succès.
   - Le reste, ce sont des bios avec des e-mails, des téléphones et des comptes de paiement. Le garder contredirait la minimisation.
4. **Pourquoi ne pas fusionner les trois échantillons ?**
   - Ils sont sélectionnés différemment : l'algorithme pour les tendances, le tri par likes pour l'humour, un choix éditorial pour les likées.
   - Leurs niveaux d'engagement ne sont pas comparables (médianes 10,7 %, 16,3 % et 18,3 %).
   - Un effet n'est jugé solide que s'il se répète dans chaque échantillon.
5. **Corrélation ou causalité ?** Ce sont des corrélations. Les abonnés sont mesurés en août 2021, après les vidéos : ils peuvent être une conséquence du succès.
6. **Que ne sait-on pas ?**
   - On ne voit aucun échec, donc on ne sait pas ce qui fait entrer une vidéo dans les tendances.
   - On ne connaît pas la situation actuelle : les données datent de 2021.
   - On ne sait rien de l'audience (âge, pays), du temps de visionnage, des ventes ni du coût des créateurs.
   - On n'a pas le contenu réel des vidéos.
7. **Que veut dire « p » ?** C'est la probabilité d'observer un tel écart par hasard. Je retiens un effet si p < 0,05, c'est-à-dire moins d'une chance sur 20.
8. **Pourquoi la durée n'est-elle qu'une piste ?**
   - Elle n'est significative que dans un échantillon sur trois.
   - Elle repose sur seulement 17 vidéos de plus de 15 s dans les tendances.
   - Sur une quarantaine de tests, environ 2 faux positifs sont attendus.
9. **Comment sont définis les sujets ? Est-ce fiable ?**
   - Par des mots-clés génériques (dog, kids, dance, funny…), avec un ordre de priorité, puis relus à la main.
   - C'est approximatif : 37 vidéos tendance n'ont pas de sujet déclaré.
   - Ce n'est jamais présenté comme un facteur prouvé.
10. **Comment sont choisis les leaders ? Pourquoi BTS ?**
    - Meilleur indice d'engagement à portée au moins médiane.
    - BTS a 37 % d'engagement, soit 3,5 fois la médiane. C'est un fait des données.
    - L'onglet signale que certains leaders sont des célébrités ou des institutions, avec des conditions sans rapport avec celles des créateurs.
11. **Pourquoi nommer certains créateurs et pas les autres ?**
    - Le client a besoin de savoir qui contacter, donc seuls les recommandés sont nommés (leader d'engagement ou de portée de chaque sujet).
    - Les autres n'apportent rien à la décision : ils sont pseudonymisés.
12. **L'agence a-t-elle le droit de détenir ces données ?**
    - Oui, sur la base de l'intérêt légitime : données rendues publiques, finalité limitée (étude pour une campagne), minimisation, durée limitée, pas de décision automatisée.
    - Les données brutes ne sont pas versionnées dans le dépôt.
    - Les bios et légendes ne sont même pas conservées dans la table de travail.
13. **Un créateur demande la suppression : que fait-on ?**
    - On supprime ses lignes des fichiers de l'agence et on prévient le client, qui supprime sa copie (art. 17 et 21).
    - On régénère le classeur avec les scripts : c'est reproductible.
14. **Comment être sûr que rien d'exclu n'est récupérable ?**
    - Le script d'audit contrôle :
      - les feuilles et lignes masquées ;
      - les noms définis ;
      - les caches de TCD (champs autorisés seulement) ;
      - les liens et les macros ;
      - les métadonnées (auteur, chemin local, étiquette de confidentialité) ;
      - la présence de chacun des 741 identifiants, 633 textes libres et 288 pseudos non retenus.
    - Résultat : tout est PASS (sauf le TreeMap tant qu'il n'est pas inséré).
15. **Les chiffres sont-ils fiables ?**
    - Ce sont des formules Excel vivantes sur l'onglet Données.
    - Excel recalcule exactement les mêmes valeurs que Python sur les 5 768 cellules numériques (script `05_verify_numbers.py --excel`).
    - Seules les p-values des tests de Mann-Whitney et Kruskal-Wallis, et deux corrélations, viennent de Python (Excel n'a pas ces tests).

## Trois pièges et comment répondre honnêtement

1. **« Donc les petits comptes sont meilleurs ? »** Non, pas forcément. On ne voit que les petits comptes qui ont percé (biais du survivant). Ce qu'on peut dire : *une fois percé*, la taille du compte ne garantit pas plus de portée, et l'engagement par vue est plus élevé. Le bon usage : sélectionner sur l'engagement mesuré et négocier les tarifs en conséquence.
2. **« Les vidéos de plus de 15 s augmentent l'engagement de 5,5 points. »** Non : c'est observé dans un seul échantillon, sur 17 vidéos, et c'est le seul test significatif parmi une quarantaine. Le verdict honnête est « effet isolé, à tester en A/B ».
3. **« L'humour a plus d'engagement (16 %) que les tendances (11 %), il faut faire de l'humour. »** Ces deux chiffres ne sont pas comparables. L'échantillon humour est le top 100 trié par likes : son engagement est gonflé par construction. Dans les tendances, l'humour est même le sujet le moins engageant (7,9 %). C'est pour ça que les seuils et les comparaisons restent internes à chaque échantillon.
