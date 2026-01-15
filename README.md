# DocuGen Pro v2.5

Generateur de courriers Word automatise - Application desktop hors ligne.

## Objectif

Generer automatiquement des courriers Word (.docx) a partir de:
- Un modele Word contenant des balises `{{ }}`
- Un fichier Excel ou chaque ligne correspond a un courrier

### Regle cle

Le logiciel **ne modifie JAMAIS le document original**. Il remplace uniquement les balises lors de la generation, en conservant totalement la mise en forme.

## Fonctionnalites

### Protection du document original
- **Lecture seule** : Le modele Word n'est jamais modifie
- Previsualisation avec tous les styles preserves
- Generation directe sans alteration du template

### Controle des sauts de page
- **Checkbox dedié** : "Ajouter un saut de page entre chaque courrier"
  - **Coche** (par defaut) : Ajoute des sauts de page automatiquement
  - **Decoche** : Si votre modele gere deja les sauts de page
- Controle total sur la pagination du document final

### Interface moderne
- Navigation par etapes (Modele → Excel → Generation)
- **Boutons ameliores** avec styles professionnels (padding, font-weight, hover)
- Previsualisation riche du modele avec surlignage des balises
- Tableau Excel editable en temps reel
- Barre de progression detaillee
- Design epure sans emojis

### Conservation totale
- Tableaux avec tous leurs styles
- En-tetes et pieds de page
- Polices, tailles, couleurs
- Gras, italique, souligne
- Alignement des paragraphes

## Installation

### Prerequis
- Python 3.8+
- Windows 10/11

### Dependances
```bash
pip install -r requirements.txt
```

### Lancement
```bash
python main.py
```

## Utilisation

### Etape 1: Modele Word

Creez un document Word avec des balises:
```
Madame, Monsieur {{nom}},

Nous avons le plaisir de vous informer concernant {{sujet}}.

Cordialement,
{{signataire}}
```

### Etape 2: Fichier Excel

| nom | sujet | signataire |
|-----|-------|------------|
| Martin | Formation | Jean Dupont |
| Dubois | Inscription | Marie Claire |

### Etape 3: Generation

1. **Selectionnez le modele Word**
   - Visualisez le contenu avec styles preserves
   - Verifiez les balises detectees

2. **Importez le fichier Excel**
   - Consultez les donnees en tableau
   - Modifiez si necessaire (double-clic)

3. **Configurez les sauts de page**
   - Cochez pour ajouter des sauts automatiques
   - Decochez si votre modele les gere deja

4. **Generez les courriers**
   - Le document original reste intact
   - Tous les styles sont preserves

## Balises disponibles

### Colonnes Excel
Toutes les colonnes deviennent des balises:
- `{{nom}}`
- `{{prenom}}`
- `{{email}}`
- etc.

### Balises automatiques
Basees sur la colonne `Sexe`:
- `{{civ_titre}}` → "Monsieur le Directeur" / "Madame la Directrice"
- `{{politesse_courte}}` → "Monsieur" / "Madame"
- `{{formule_politesse_fin}}` → Formule complete

## Architecture

```
courrier_gen/
├── main.py                 # Point d'entree
├── ui/
│   └── main_window.py      # Interface + WordPreview
├── logic/
│   ├── processor.py        # Moteur de generation
│   └── rules.py           # Regles metier
└── requirements.txt
```

## Technologies

- **PySide6**: Interface graphique Qt6
- **python-docx**: Lecture des documents Word
- **docxtpl**: Remplacement des balises (Jinja2)
- **pandas**: Manipulation des donnees Excel
- **openpyxl**: Lecture/ecriture Excel

## Points cles

### Document original protege
- **JAMAIS modifie** pendant le processus
- Generation directe sans alteration
- Previsualisation en lecture seule

### Controle des sauts de page
- **Checkbox explicite** pour le controle utilisateur
- Adapte selon vos besoins
- Evite les doublons et les sauts inutiles

### Styles des boutons
- **Padding ameliore** : 10-16px vertical, 24-32px horizontal
- **Font-weight: 600/bold** pour meilleure lisibilite
- **Hover effects** : Assombrissement au survol
- **Border-radius: 8-12px** pour design moderne

## Depannage

### Les balises ne sont pas remplacees
- Verifiez que les noms de colonnes Excel correspondent exactement aux balises
- Exemple: `{{nom}}` necessite une colonne "nom" (sensible a la casse)

### Les styles sont perdus
- Verifiez que c'est bien un fichier .docx (pas .doc)
- Le document original n'est jamais modifie, donc les styles sont garantis

### Sauts de page non desires
- **Decochez** "Ajouter un saut de page entre chaque courrier"
- Utilisez cette option si votre modele gere deja les sauts

### Sauts de page manquants
- **Cochez** "Ajouter un saut de page entre chaque courrier"
- Un saut sera ajoute automatiquement entre chaque courrier

## Performances

- Vitesse: ~10-50 courriers/seconde
- Capacite: Teste avec 500+ courriers
- Memoire: Optimise avec streaming

---

**Version**: 2.5
**Licence**: MIT
**Developpe avec**: Python | Qt | Protection du document original
