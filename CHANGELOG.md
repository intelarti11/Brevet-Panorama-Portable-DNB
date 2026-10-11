# Historique

## 0.1.4

- Retrait de l’INE et du statut boursier individuel des rapports complets PDF et XLSX. Les comparatifs agrégés restent disponibles.

## 0.1.3

- Lecture des fichiers Excel avec SheetJS CE 0.20.3, corrigé pour
  CVE-2023-30533 et CVE-2024-22363. Les exports conservent leurs styles.
- Rejet des textes numériques invalides et des notes hors barème dans tous
  les imports DNB, avec ou sans métadonnées du modèle. Les décimales à virgule,
  notes à zéro, absences et barèmes historiques restent pris en charge.
- Conservation du sexe des élèves lors de l’import DNB.
- Conservation des moyennes à zéro dans le graphique comparatif BB1/BB2.
