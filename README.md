# Estimation immobilière VBA

Ce projet est un outil d'estimation immobilière développé en VBA sous Excel.

## Fonctionnalités

- **Base de données DVF 2024** : Le fichier contient une base de données de biens immobiliers avec leurs caractéristiques (surface, département, nombre de pièces, type de logement, ville, etc.).
- **Formulaire de saisie** : Un UserForm permet de renseigner les critères du bien à estimer (voir capture d'écran).
    - Surface (m²)
    - Département
    - Nombre de pièces
    - Type de logement
    - Ville
- **Estimation automatique** : Après saisie des critères, le programme calcule et affiche le prix moyen correspondant, basé sur les transactions enregistrées en base.
- **Code VBA accessible** : Le code source du formulaire et des modules est inclus dans le dépôt (`UserForm1.frm`, `UserForm1.frx`).

## Utilisation

1. **Ouvrir le fichier Excel** : `PROJET VBA ANALYSE IMMOBILIERE.xlsm`
2. **Activer les macros** pour pouvoir utiliser le formulaire.
3. **Cliquer sur le bouton d'estimation** pour ouvrir le formulaire.
4. **Saisir les critères** du bien immobilier dans le formulaire.
5. **Valider** pour obtenir le prix moyen estimé en fonction des biens similaires présents dans la base 2024.

## Structure du dépôt

- `PROJET VBA ANALYSE IMMOBILIERE.xlsm` : Fichier principal Excel avec la base de données et le code VBA.
- `UserForm1.frm` & `UserForm1.frx` : Fichiers représentant le formulaire de saisie (à importer dans un projet VBA si besoin).


## Exemple d'écran

<img width="1360" height="768" alt="image" src="https://github.com/user-attachments/assets/4ee6148a-7e4f-45c0-9f32-8607736f7f21" />


## Informations

- Base DVF 2024 pour test et estimation.
- Projet VBA, compatible Excel avec macros activées.
- Contact : antoinespion@protonmail.com

---

**NB** : Pour utiliser ou modifier le UserForm dans un autre projet, importer à la fois le `.frm` et le `.frx` dans l'éditeur VBA.
