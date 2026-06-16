# DAEU Numérique — Épreuve n°3 · Session 2, 2026

> **Université du Mont-Houy · Valenciennes · Académie de Lille**
> Module Informatique · Niveau 10 · Professeur : M. Sautière

---

## Présentation de l'épreuve

L'épreuve n°3 du module Numérique du DAEU évalue deux compétences complémentaires : la **maîtrise des fonctions logiques dans Excel** et la **programmation en Python**. Elle se déroule en **2 heures** et est notée sur **20 points**.

| Partie | Format | Compétences évaluées | Durée conseillée | Barème |
|--------|--------|----------------------|------------------|--------|
| **Partie 1 — QCM** | Papier | Logique Excel · Python (variables, conditions, boucles, listes) | ≤ 75 min | 20 pts |
| **Partie 2 — Excel** | Ordinateur | Fonctions SI, ET, OU — formules à paliers | ≤ 45 min | 15 pts |
| | | **Total ramené sur 20** | **2h** | **35 → /20** |

> **Ordre obligatoire** : le QCM papier doit être rendu au professeur **avant** d'allumer les ordinateurs. Aucun retour en arrière possible.

---

## 📁 Fichiers du dépôt

```
pratique_2026/
├── README.md                            ← ce fichier
├── DAEU_Logique_Epreuve3_excel.xlsx     ← fichier Excel à télécharger et compléter
└── DAEU_Epreuve3_Consignes.docx         ← document de consignes complet (impression)
```

---

## Partie 2 — Exercices Excel

### Téléchargement et mise en route

1. **Allumez l'ordinateur** uniquement après avoir remis votre copie QCM au professeur.
2. Rendez-vous sur ce dépôt : **[darksathili-jpg.github.io/pratique_2026](https://darksathili-jpg.github.io/pratique_2026)**
3. Téléchargez le fichier **`DAEU_Logique_Epreuve3_excel.xlsx`** et ouvrez-le dans Excel.
4. Enregistrez-le immédiatement sous **`NOM_Prénom_Excel.xlsx`** sur le bureau.
5. À la fin de l'épreuve, déposez votre fichier dans le répertoire portant votre nom sur la clé USB mise à disposition.

> 💡 Ne fermez pas Excel avant que le professeur ait confirmé la récupération de votre fichier.

---

### Structure du fichier Excel

Le fichier contient **trois feuilles de calcul**, une par exercice. Chaque feuille présente :
- une colonne **orange** → à compléter avec votre formule
- une colonne **verte** → résultat attendu (ne pas modifier)

---

### Exercice 1 — Licences de football `SI + ET` · 5 pts

**Contexte :** Le club US Valenciennois délivre une licence officielle uniquement si le joueur remplit **toutes** les conditions simultanément.

**Consigne — colonne F :**
```excel
=IF(ET(...)) → "Licence accordée"  ou  "Dossier incomplet"
```

**Critères :**
- Âge `>= 8` ans
- Certificat médical `= "Oui"`
- Cotisation `= "Payée"`

**Barème :**

| Critère d'évaluation | Points |
|----------------------|--------|
| Formule `IF(ET(...))` syntaxiquement correcte | 1 pt |
| 3 critères correctement codés (0,5 pt chacun) | 1,5 pt |
| 6 résultats exacts (0,25 pt chacun) | 1,5 pt |
| Guillemets et séparateurs conformes | 1 pt |

---

### Exercice 2 — Prêt prolongé à la bibliothèque `SI + OU` · 5 pts

**Contexte :** La médiathèque de Valenciennes accorde un prêt de **28 jours** (au lieu de 14) si **au moins une** condition est vraie.

**Consigne — colonne F :**
```excel
=IF(OU(...)) → "Prêt 28 jours"  ou  "Prêt standard 14 jours"
```

**Critères :**
- Carte senior `= "Oui"`
- OU Statut `= "Étudiant"`

**Barème :**

| Critère d'évaluation | Points |
|----------------------|--------|
| Formule `IF(OU(...))` syntaxiquement correcte | 1 pt |
| 2 critères correctement codés (0,5 pt chacun) | 1 pt |
| 6 résultats exacts (0,33 pt chacun) | 2 pts |
| Messages affichés conformes à la consigne | 1 pt |

---

### Exercice 3 — Frais de livraison traiteur `SI + ET + OU imbriqués` · 5 pts

**Contexte :** Le traiteur *Saveurs du Nord* applique des frais de livraison à **paliers** selon le montant et le nombre de couverts.

**Consigne — colonne F :**
```excel
=IF(OU(montant>=80 ; couverts>=10) ; "GRATUIT" ;
   IF(couverts>=5 ; "RÉDUIT (3,50 €)" ;
      "STANDARD (7,00 €)"))
```
> **Indice :** tester la gratuité en **premier** avec `SI + OU`, puis imbriquer les `SI` pour les paliers suivants.

**Barème :**

| Critère d'évaluation | Points |
|----------------------|--------|
| Structure d'imbrication `SI` correcte | 1 pt |
| Test de gratuité avec `OU` placé en premier | 1,5 pt |
| Palier RÉDUIT correctement codé | 1 pt |
| 6 résultats exacts (0,25 pt chacun) | 1,5 pt |

---

## Partie 1 — QCM (rappel de structure)

Le QCM est distribué en **format papier** et n'est pas disponible sur ce dépôt. Il comporte **40 questions** réparties en 10 parties.

| Partie | Thème | Questions | Barème |
|--------|-------|-----------|--------|
| 1 | Valeurs logiques & opérateurs de comparaison | Q1 – Q3 | 3 pts |
| 2 | Fonction SI (simple et imbriqué) | Q4 – Q7 | 4 pts |
| 3 | Fonctions ET, OU, NON | Q8 – Q11 | 4 pts |
| 4 | Combinaisons SI + ET / SI + OU | Q12 – Q15 | 4 pts |
| 5 | Variables, types et saisie utilisateur | Q16 – Q20 | 5 pts |
| 6 | Structures conditionnelles | Q21 – Q25 | 5 pts |
| 7 | Boucles `for` / `while` | Q26 – Q30 | 5 pts |
| 8–10 | Listes et parcours | Q31 – Q40 | 10 pts |

**Modalités :** une seule case à cocher par question · en cas d'erreur barrer et entourer · Q5 et Q12 nécessitent une formule rédigée.

---

## 🎯 Compétences évaluées

```
Excel                              Python
──────────────────────────────     ──────────────────────────────
✔ Valeurs logiques VRAI/FAUX       ✔ Variables et types (int, str, float, bool)
✔ Opérateurs de comparaison        ✔ Saisie utilisateur (input / int())
✔ Fonction SI simple               ✔ Structures if / elif / else
✔ Fonction SI imbriqué             ✔ Boucles for et while
✔ Fonctions ET, OU, NON            ✔ Listes : création, accès, modification
✔ Combinaisons SI + ET / SI + OU   ✔ Parcours de listes et compréhension
✔ Formules à paliers multiples     ✔ Algorithmes de recherche et d'accumulation
```

---

## Informations pratiques

| | |
|-|-|
| **Établissement** | Lycée du Mont-Houy, Valenciennes |
| **Formation** | DAEU — Diplôme d'Accès aux Études Universitaires |
| **Module** | Informatique · Niveau 10 |
| **Session** | Session 2 — Juin 2026 |
| **Date** | 25 / 06 / 2026 |
| **Durée** | 2 heures |
| **Professeur** | M. Sautière |
| **Documents autorisés** | Aucun |

---

## 🔗 Ressources du cours

- Dépôt principal du cours : [darksathili-jpg.github.io/pratique_2026](https://darksathili-jpg.github.io/pratique_2026)
- Feuille de cours Excel (logique) : disponible sur le dépôt principal
- Supports Python : disponibles sur le dépôt principal

---

<sub>Épreuve n°3 · DAEU Numérique · Session 2, 2026 · Lycée du Mont-Houy, Valenciennes · M. Sautière</sub>
