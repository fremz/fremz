# Calculateur de Ratios Financiers — PME Belges
## Fidunot Expertise SRL — Usage confidentiel local

---

## Installation (une seule fois)

### Prérequis
- Python 3.9 ou supérieur : https://www.python.org/downloads/
  → Cocher "Add Python to PATH" lors de l'installation

### Installation des dépendances
Ouvrir un terminal (cmd ou PowerShell) dans le dossier du projet :

```
pip install -r requirements.txt
```

---

## Lancement

Double-cliquer sur `lancer.bat`  
ou depuis un terminal :

```
python ratios_financiers.py
```

---

## Conversion en .exe Windows (optionnel)

Pour distribuer l'outil sans installer Python sur chaque poste :

```
pip install pyinstaller
pyinstaller --onefile --windowed --name "RatiosFinanciers" ratios_financiers.py
```

Le fichier `.exe` sera généré dans le dossier `dist/`.

---

## Utilisation

1. **Saisie des données** — Onglet "Saisie des données"
   - Renseigner le nom du client, le type et les exercices N / N-1
   - Saisir les postes comptables pour N et N-1 (références PCMN incluses)
   - Appuyer sur "Calculer" ou Tab/Entrée après chaque saisie

2. **Ratios calculés** — Onglet "Ratios calculés"
   - Affichage automatique avec statut vert / orange / rouge
   - Flèche de tendance N vs N-1
   - Seuils adaptés au type de client sélectionné

3. **Altman Z'** — Onglet "Altman Z'"
   - Cliquer "Auto-remplir" pour pré-calculer X1, X3, X4, X5 depuis les données saisies
   - Saisir X2 (réserves / total bilan) manuellement
   - Score Z' calculé avec zone de diagnostic

4. **Export Excel**
   - Cliquer "Exporter Excel"
   - Choisir le dossier d'enregistrement
   - Le fichier Excel contient : onglet Analyse (ratios formatés) + onglet Données saisies

---

## Confidentialité

Aucune donnée n'est transmise à l'extérieur.
L'application fonctionne intégralement en local, sans connexion réseau.

---

## Support

Pour toute question : cabinet Fidunot Expertise SRL
