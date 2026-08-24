# Dashboard RH 2026
## Tableau de Bord Ressources Humaines

---

## 📋 Description

Dashboard RH professionnel et interactif permettant de piloter l'ensemble des indicateurs RH de l'entreprise :
- Effectifs et pyramides (âge, ancienneté)
- Masse salariale et coûts
- Mouvements (recrutements, départs)
- Indicateurs de performance RH

---

## 🚀 Installation et Déploiement

### Structure des fichiers

```
Dashboard_RH/
├── app.py                  # Application Flask (serveur HTTPS)
├── generate_data.py        # Génération des données (EFFECTIF/MS/MOUVEMENT/FORMATION) — source de vérité
├── index.html              # Dashboard principal
├── generate_cert.py        # Générateur de certificats SSL
├── start.sh                # Script de démarrage rapide
├── requirements.txt        # Dépendances Python
├── data/
│   ├── TDB_COURANT.xlsx    # Fichier Excel source RH (à mettre à jour mensuellement)
│   ├── FORMATION.xlsx      # Fichier Excel source Formation (plan de formation, à mettre à jour en continu)
│   ├── SOCIAL.xlsx         # Fichier Excel source Social (actions sociales et sociétales)
│   └── data.json           # Données générées (automatique)
├── scripts/
│   └── generate_data.py    # Point d'entrée du cron mensuel (appelle generate_data.py à la racine)
├── cert.pem                # Certificat SSL (à générer)
├── key.pem                 # Clé privée SSL (à générer)
└── README.md               # Ce fichier
```

### Déploiement sur le serveur

**Serveur cible :** `10.222.36.9`  
**Port :** `5001`  
**URLs d'accès :**
- `https://portailrh.sonasid.ma:5001`
- `https://10.222.36.9:5001`

#### Étape 1: Copier les fichiers sur le serveur

```bash
scp -r Dashboard_RH/ user@10.222.36.9:/opt/dashboard-rh/
```

#### Étape 2: Installer les dépendances Python

```bash
cd /opt/dashboard-rh
pip3 install -r requirements.txt
```

#### Étape 3: Configurer le certificat SSL

**Option A - Certificat auto-signé (développement):**
```bash
python3 generate_cert.py
# Choisir l'option 1
```

**Option B - Certificat officiel (production):**
```bash
# Copier vos certificats
cp /chemin/vers/votre/certificat.pem cert.pem
cp /chemin/vers/votre/cle-privee.pem key.pem
chmod 600 key.pem
chmod 644 cert.pem
```

**Option C - Convertir un fichier PFX:**
```bash
python3 generate_cert.py
# Choisir l'option 2 et fournir le fichier PFX
```

#### Étape 4: Démarrer le serveur

```bash
# Méthode simple
./start.sh

# Ou directement
python3 app.py
```

#### Étape 5: Configurer en service (production)

Créer `/etc/systemd/system/dashboard-rh.service`:

```ini
[Unit]
Description=Dashboard RH 2026
After=network.target

[Service]
Type=simple
User=www-data
WorkingDirectory=/opt/dashboard-rh
ExecStart=/usr/bin/python3 /opt/dashboard-rh/app.py
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
```

Activer le service:
```bash
sudo systemctl daemon-reload
sudo systemctl enable dashboard-rh
sudo systemctl start dashboard-rh
sudo systemctl status dashboard-rh
```

---

## 📊 Mise à jour des données

### Procédure mensuelle

1. **Déposer le nouveau fichier Excel** dans le dossier `data/` :
   - Renommer le fichier en `TDB_COURANT.xlsx`
   - Ou modifier la variable `DATA_FILE` dans `scripts/generate_data.py`

2. **Exécuter le script de génération** :
   ```bash
   cd /var/www/portailrh/Agirh/Dashboard
   python3 scripts/generate_data.py
   ```

3. **Vérifier la mise à jour** :
   - Le fichier `data/data.json` est mis à jour
   - Actualiser le navigateur pour voir les nouvelles données

### Prérequis Python

```bash
pip install pandas openpyxl numpy
```

### Structure du fichier Excel attendu (TDB_COURANT.xlsx)

Le fichier Excel doit contenir 3 feuilles :

1. **EFFECTIF** - Liste des collaborateurs
   - MATRICULE, NOM, PRENOM, etb_name, CLASSIFICATION_LIB
   - DT_NAISSANCE, DEB_ADMINISTRATION, TYPE_CONTRAT
   - CIVILITE, FONCTION_LIB, HIERARCHIE_LIB, etc.

2. **MOUVEMENT** - Recrutements et départs
   - Section "État des recrutements" : Matricule, Nom, Prénom, Site, Collège, Date, Fonction
   - Section "État des départs" : Matricule, Nom, Prénom, Site, Collège, Date, Motif

3. **MS** - Masse Salariale
   - COMPTE_IMPUTATION, INTITULE_COMPLET, etb_name, Classification
   - Colonnes mensuelles : Janvier, Février, ..., Décembre

### Structure du fichier Formation attendu (data/FORMATION.xlsx)

Fichier optionnel : s'il est absent, la page **Formation** affiche un état vide.
Une seule feuille, une ligne par participant, colonnes dans cet ordre : N° Action,
Domaine, Type d'action, Thèmatique, Date Début, Date Fin, Nbre de Jours, Nbre
d'heure, Matricule, Nom, Prénom, Classification, Site, Département, Statut,
Cabinet de formation, Nom Formateur (interne), Coût Formation, Coût Logistique,
Evaluation à chaud, Evaluation Formateur, Taux de Satisfaction, Evaluation à
froid, Résultat. Les colonnes propres à l'action (Domaine, dates, coûts...) ne
sont attendues que sur la première ligne de chaque formation ; le script
propage automatiquement ces valeurs sur les lignes suivantes.

### Structure du fichier Social/Sociétal attendu (data/SOCIAL.xlsx)

Fichier optionnel : s'il est absent, les pages **Social** et **Sociétal**
affichent un état vide. Une seule feuille, une ligne par action, colonnes
dans cet ordre : Actions (description), Type Action (`Social` ou
`Sociétales` — c'est cette colonne qui répartit chaque ligne vers la page
Social ou la page Sociétal), Date de réalisation (texte libre du type
"Février 2026" ou "Juin - Juillet 2026"), Nombre de bénéficiaires, Budget,
Région, Site(s) Sonasid concerné(s). Les régions/sites multiples sont
séparés par des virgules.

Les budgets *alloués* (par opposition aux budgets *consommés*, calculés
depuis les fichiers Excel) se saisissent tous depuis la page **Budget**,
protégée par un mot de passe administrateur, et restent stockés dans le
navigateur (pas de rechargement Excel nécessaire pour les ajuster).

---

## 📈 Indicateurs calculés

### Indicateurs de Rotation
| Indicateur | Formule |
|------------|---------|
| **Turnover** | (Recrutements + Départs) / 2 / FTE base × 100 |
| **Taux de démission** | Démissions / FTE base × 100 |
| **Taux de rétention** | 100% - Taux de départ |
| **Taux de croissance** | Solde net / FTE base × 100 |
| **Réussite période essai** | % recrutements validés |

### Indicateurs Démographiques
| Indicateur | Description |
|------------|-------------|
| **Taux de féminisation** | Femmes / Total × 100 |
| **Âge moyen** | Moyenne d'âge des collaborateurs |
| **Indice de jeunesse** | % collaborateurs < 35 ans |
| **Indice de vieillissement** | % collaborateurs > 50 ans |

### Indicateurs de Stabilité
| Indicateur | Description |
|------------|-------------|
| **Indice de stabilité** | % ancienneté ≥ 5 ans |
| **Taux de CDI** | CDI / Total × 100 |
| **Ratio d'encadrement** | Cadres / Total × 100 |

### Indicateurs Financiers
| Indicateur | Description |
|------------|-------------|
| **Masse salariale YTD** | Cumul depuis début d'année |
| **Coût par ETP** | MS mensuelle / FTE |
| **Budget annuel estimé** | Projection sur 12 mois |

### Indicateurs Formation
| Indicateur | Description |
|------------|-------------|
| **Taux de couverture** | Collaborateurs formés (distincts) / Effectif total × 100 |
| **Taux de présence** | Présents / (Présents + Absents) × 100 |
| **Heures Formation** | Somme des heures de formation × participants |
| **Coût moyen par participant** | (Coût formation + logistique) / Nb participations |
| **Taux d'investissement formation** | Coût total formation / Masse salariale YTD × 100 |
| **Synthèse par cabinet** | Actions/participations/jours/coût regroupés par cabinet de formation, filtrable par année |

### Indicateurs Social / Sociétal / Budget
| Indicateur | Description |
|------------|-------------|
| **Budget consommé** | Somme des budgets des actions réalisées, par catégorie (Social, Sociétal, Formation, Masse Salariale) |
| **Budget FTE** | Coût annuel cible par collaborateur, comparé au coût réel (Coût par ETP mensuel × 12) |
| **Budget alloué** | Saisi par un administrateur sur la page Budget (mot de passe requis), non issu d'un fichier |
| **Taux de réalisation** | Budget consommé / Budget alloué × 100, par catégorie |

---

## 🎨 Fonctionnalités

- ✅ **9 pages** : Synthèse, Effectifs, Masse Salariale, Mouvements, Indicateurs, Formation, Social, Sociétal, Budget
- ✅ **Filtres dynamiques** : Établissement, Classification, Type de contrat
- ✅ **Recherche globale** : Matricule, Nom, Prénom, Fonction
- ✅ **Export CSV** : Liste des collaborateurs
- ✅ **Graphiques interactifs** : Chart.js 4.4
- ✅ **Design responsive** : Adapté mobile/tablette
- ✅ **Thème clair** : Interface épurée et professionnelle

---

## 🔧 Configuration serveur

### Apache (exemple .htaccess)

```apache
<IfModule mod_headers.c>
    Header set Cache-Control "no-cache, no-store, must-revalidate"
    Header set Pragma "no-cache"
    Header set Expires 0
</IfModule>

<FilesMatch "\.(json)$">
    Header set Content-Type "application/json; charset=utf-8"
</FilesMatch>
```

### Nginx (exemple configuration)

```nginx
location /Agirh/Dashboard {
    root /var/www/portailrh;
    index index.html;
    
    location ~* \.json$ {
        add_header Content-Type "application/json; charset=utf-8";
        add_header Cache-Control "no-cache";
    }
}
```

---

## 📞 Support

Pour toute question ou problème technique, contactez le département RH ou l'équipe IT.

---

**Version :** 2.0  
**Date :** Janvier 2026  
**Auteur :** Département RH
