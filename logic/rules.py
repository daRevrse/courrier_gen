def get_civility_rules(row):
    """Logique métier pour les titres et formules de politesse."""
    sexe = str(row.get('Sexe', '')).strip().lower()
    fonction = str(row.get('Fonction', '')).strip()
    
    # Détermination du préfixe et du titre
    if sexe in ['homme', 'h', 'm', 'monsieur']:
        prefixe = "Monsieur"
        titre_complet = f"Monsieur le {fonction}" if fonction else "Monsieur"
    else:
        prefixe = "Madame"
        titre_complet = f"Madame la {fonction}" if fonction else "Madame"
    
    # Formule finale
    formule_fin = f"Veuillez agréer, {prefixe}, l'expression de mes salutations distinguées."
    
    return {
        "civ_titre": titre_complet,
        "politesse_courte": prefixe,
        "formule_politesse_fin": formule_fin
    }