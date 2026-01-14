import os
import re
import pandas as pd
from docx import Document
from docxtpl import DocxTemplate
from io import BytesIO
from .rules import get_civility_rules

class MailGenerator:
    def __init__(self, template_path, excel_path, output_dir):
        """
        Initialise le moteur de génération.
        :param template_path: Chemin du fichier Word (.docx)
        :param excel_path: Chemin du fichier Excel (.xlsx)
        :param output_dir: Dossier où sera enregistré le fichier final
        """
        self.template_path = template_path
        self.excel_path = excel_path
        self.output_dir = output_dir
        # Balises gérées automatiquement par le moteur de règles (logic/rules.py)
        self.auto_tags = ['civ_titre', 'politesse_courte', 'formule_politesse_fin']

    @staticmethod
    def extract_tags(template_path):
        """
        Analyse le document Word pour extraire les balises {{variable}}.
        Utile pour informer l'utilisateur des colonnes Excel nécessaires.
        """
        doc = Document(template_path)
        full_text = []
        
        # Extraction du texte des paragraphes
        for para in doc.paragraphs:
            full_text.append(para.text)
        
        # Extraction du texte des tableaux
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    full_text.append(cell.text)
        
        combined_text = " ".join(full_text)
        # Trouve tout ce qui est entre {{ }}
        found_tags = re.findall(r"\{\{\s*(\w+)\s*\}\}", combined_text)
        
        # On ignore les balises calculées automatiquement
        auto_tags = ['civ_titre', 'politesse_courte', 'formule_politesse_fin']
        return sorted(list(set(t for t in found_tags if t not in auto_tags)))

    def validate_excel(self):
        """
        Vérifie la correspondance entre les balises du Word et les colonnes Excel.
        """
        needed_tags = self.extract_tags(self.template_path)
        df = pd.read_excel(self.excel_path)
        
        # Nettoyage des noms de colonnes (suppression des espaces vides)
        df.columns = [str(c).strip() for c in df.columns]
        actual_columns = df.columns.tolist()
        
        missing = [tag for tag in needed_tags if tag not in actual_columns]
        return missing, df

    def run(self, progress_callback):
        """
        Exécute la génération groupée dans un seul fichier.
        :param progress_callback: Fonction pour mettre à jour la barre de progression UI
        """
        # 1. Validation de sécurité
        missing, df = self.validate_excel()
        if missing:
            raise ValueError(f"Erreur : Les colonnes suivantes manquent dans l'Excel : {', '.join(missing)}")

        total = len(df)
        master_doc = None

        # 2. Boucle de génération
        for index, row in df.iterrows():
            # Génération du courrier individuel en mémoire vive (RAM)
            tpl = DocxTemplate(self.template_path)
            
            # Préparation des données (Données Excel + Règles métier)
            context = row.to_dict()
            context.update(get_civility_rules(row))
            
            # Rendu du template (remplacement des balises)
            tpl.render(context)
            
            # Sauvegarde temporaire dans un flux binaire
            temp_stream = BytesIO()
            tpl.save(temp_stream)
            temp_stream.seek(0)
            
            # Chargement du document généré pour fusion
            current_doc = Document(temp_stream)
            
            if master_doc is None:
                # Le premier courrier devient la base du document maître
                master_doc = current_doc
            else:
                # Ajout d'un saut de page avant d'ajouter le courrier suivant
                master_doc.add_page_break()
                
                # Fusion des éléments (paragraphes et tableaux) dans le document maître
                for element in current_doc.element.body:
                    # On évite de copier les propriétés de section pour ne pas casser le saut de page
                    if element.tag.endswith('sectPr'):
                        continue
                    master_doc.element.body.append(element)

            # Mise à jour de l'interface
            progress_callback(index + 1, total)

        # 3. Sauvegarde finale du document groupé
        if master_doc:
            final_filename = "Impression_Groupée_Courriers.docx"
            output_path = os.path.join(self.output_dir, final_filename)
            master_doc.save(output_path)
            return output_path
        else:
            raise Exception("Aucune donnée n'a été traitée.")