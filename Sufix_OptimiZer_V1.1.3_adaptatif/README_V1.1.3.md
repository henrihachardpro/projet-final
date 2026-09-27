# SUFIX OptimiZer V1.1.3

Évolutions principales :
- Espacement entre supports extrait du PDF et validé dans l'interface.
- Châssis : nombre proposé à 1, espacement verrouillé.
- Chemins de câbles/couvercles : calcul métrique, barres commerciales de 3 m via Packaging, feuille Opti Chemin de câbles.
- Interface assistant plein écran, minimisable, scroll, retour/suivant, glisser-déposer.
- Détection automatique PDF associé.
- Contrôles de cohérence + feuille Contrôles.
- Version des bases tracée dans data_versions.json et _Paramètres.
- Comparaison/KPI optimisation + Synthèse optimisation et KPI en Plan de découpe.
- Modes Économie matière / Équilibré / Simplicité de fabrication.
- Écran final avec nouvelle étude / fermer / exports EQ.
- Refactor : sufix_core.py + sufix_ui.py + bootstrap.
- Tests unitaires intégrés et exécutés avant compilation.

Limites de cette première V1.1.3 :
- Sans nouveaux PDF représentatifs contenant « Espacement entre support », l'extraction a été développée avec plusieurs formats textuels mais doit être validée sur les prochains projets réels.
- Le mode « Simplicité de fabrication » optimise parmi les stratégies internes existantes ; il ne s'agit pas encore d'un solveur OR-Tools exhaustif.
- Les bases restent volontairement embarquées dans l'EXE.
