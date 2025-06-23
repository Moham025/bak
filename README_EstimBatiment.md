# EstimBatiment - Guide d'utilisation

## Vue d'ensemble
La fonctionnalité **EstimBatiment** permet de générer automatiquement un devis détaillé à partir d'un fichier Excel d'estimation de bâtiment.

## Structure du fichier Excel requis

Le fichier Excel d'entrée (comme `EstimType.xlsx`) doit contenir les feuilles suivantes :

### Feuilles obligatoires :
- **`qt`** - Données de quantités (dimensions, surfaces, etc.)
- **`calcul`** - Calculs principaux (Blocs I, II, III)

### Feuilles optionnelles :
- **`open`** - Menuiserie (Bloc IV)
- **`Electricite`** - Électricité (Bloc V)
- **`Plomberie`** - Plomberie sanitaire (Bloc VI)
- **`Revetement`** - Revêtement (Bloc VII)
- **`Peinture`** - Peinture (Bloc VIII)
- **`Toiture`** - Toiture (Bloc IX)

## Comment utiliser la fonctionnalité

### 1. Démarrer le serveur local
```bash
python main.py
```
Le serveur sera accessible sur `http://localhost:5000` ou `http://192.168.11.107:5000`

### 2. Utiliser l'interface web
1. Ouvrez l'application Flutter
2. Cliquez sur le bouton **"Estim Batiment"** (icône calculatrice)
3. Sélectionnez votre fichier Excel d'estimation
4. Cliquez sur **"Lancer l'estimation"**
5. Une fois le traitement terminé, cliquez sur **"Télécharger"**

### 3. Résultat
Un fichier Excel sera généré avec :
- Tous les blocs d'estimation formatés
- Calculs automatiques des quantités
- Prix unitaires appliqués
- Récapitulatif avec total général
- Montant en lettres (en Francs CFA)

## Format de sortie

Le fichier résultat contient une feuille "Estimation Globale" avec :

- **Bloc I-III** : Terrassement, Maçonnerie, etc. (basé sur la feuille `calcul`)
- **Bloc IV** : Menuiserie (basé sur la feuille `open`)
- **Bloc V** : Électricité (basé sur la feuille `Electricite`)
- **Bloc VI** : Plomberie sanitaire (basé sur la feuille `Plomberie`)
- **Bloc VII** : Revêtement (basé sur la feuille `Revetement`)
- **Bloc VIII** : Peinture (basé sur la feuille `Peinture`)
- **Bloc IX** : Toiture (basé sur la feuille `Toiture`)
- **Récapitulatif** : Total général avec montant en lettres

## Dépannage

### Erreurs courantes :
1. **"Feuille 'qt' manquante"** : Assurez-vous que votre fichier contient une feuille nommée exactement `qt`
2. **"Feuille 'calcul' manquante"** : Assurez-vous que votre fichier contient une feuille nommée exactement `calcul`
3. **"Erreur de connexion"** : Vérifiez que le serveur Flask est démarré avec `python main.py`

### Test en ligne de commande :
```bash
python test_estim_batiment.py
```

## Structure technique

### Modules Python utilisés :
- `data_reader.py` - Lecture des données Excel
- `calculation_engine.py` - Moteur de calcul et formules
- `excel_writer.py` - Génération des tableaux Excel
- `number_to_letter_converter.py` - Conversion montant en lettres

### Endpoint API :
- **POST** `/estim-batiment`
- **Content-Type** : `multipart/form-data`
- **Paramètre** : `excel_file` (fichier Excel)
- **Réponse** : Fichier Excel généré

## Exemples de formules supportées

Dans les feuilles de calcul, vous pouvez utiliser des références comme :
- `SEMELLE[l] * SEMELLE[L] * SEMELLE[h]` - Volume de semelle
- `MUR[longueur] * MUR[hauteur]` - Surface de mur
- `LONGRINE[ml] * 2` - Longueur doublée

Les références sont automatiquement résolues à partir de la feuille `qt`. 