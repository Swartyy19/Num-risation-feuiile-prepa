# Feuille Prépa de perçage

`Feuille_Prepa_testttt.xlsm` est un classeur Excel contenant des macros VBA permettant de préparer visuellement des opérations de perçage et de taraudage. Les différents UserForms servent à sélectionner le profil et le type d'opération pour chaque côté de la pièce.

## Version d'Excel requise

Les macros ont été testées avec Excel 2013 et versions ultérieures. Une version d'Excel prenant en charge VBA est indispensable.

Lors de l'ouverture du classeur, Excel peut afficher un message de sécurité indiquant que des macros sont désactivées. Cliquer sur **Activer le contenu** pour permettre l'exécution des macros.

## Principes de fonctionnement

1. Sur la feuille **"Prpa Numrise"**, choisir le profil de la pièce dans la cellule `AL7` (par ex. `40x40L`, `45x45_2NVS`, etc.).
2. Selon le profil choisi, cliquer sur les boutons de gauche ou de droite (ex. `Bouton_G1`, `Bouton_D1`, …) pour ouvrir la fenêtre de sélection du type de perçage.
3. Dans la fenêtre qui apparaît, sélectionner le type d'opération désiré (perçage, taraudage, etc.) puis valider. Les formes correspondantes sont automatiquement affichées ou masquées sur la feuille.
4. Les boutons « Taraudage » ouvrent `UserForm_Taraudage` pour choisir le type de taraudage (T1, T2, T3).
5. Pour certains profils comme `45x45_2NVS`, le formulaire `UserForm_2NVS` permet de choisir la vue A ou B.

## Réimporter les modules VBA

Si le classeur perd ses macros, ouvrir l'éditeur VBA (`Alt` + `F11`) puis utiliser **Fichier → Importer un fichier…** pour chacun des fichiers présents dans ce dépôt :

* `AfficherFormulaire.bas`
* `Module2.bas`
* `Module6.bas`
* `taraudage.bas`
* `UserForm_ChoixPercage.frm`
* `UserForm_Doublebarre.frm`
* `UserForm_Taraudage.frm`
* `UserForm_2NVS.frm`
* `vba feuil1.cls`

Ces fichiers recréent l'ensemble des modules et UserForms nécessaires au fonctionnement de `Feuille_Prepa_testttt.xlsm`.
