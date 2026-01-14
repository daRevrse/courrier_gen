import json
import os
import subprocess
import pandas as pd
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLabel, QFileDialog, QProgressBar, 
                             QMessageBox, QFrame, QTableWidget, QTableWidgetItem, 
                             QHeaderView, QDialog, QTextBrowser, QComboBox)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from logic.processor import MailGenerator

CONFIG_FILE = "config_session.json"

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DocuGen Pro - Générateur de Courriers")
        self.setMinimumSize(950, 800)
        self.template_path = None
        self.excel_path = None
        self.last_generated_path = None # Pour le bouton "Ouvrir"
        
        self.setup_ui()
        self.load_session()

    def setup_ui(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #f8f9fa; }
            QFrame#card { background: white; border-radius: 10px; border: 1px solid #dee2e6; }
            QPushButton#primaryBtn { background: #0d6efd; color: white; font-weight: bold; border-radius: 6px; }
            QPushButton#primaryBtn:hover { background: #0b5ed7; }
            QPushButton#openBtn { background: #198754; color: white; border-radius: 6px; font-weight: bold; }
            QPushButton#openBtn:hover { background: #157347; }
            QPushButton#helpBtn { background: #0d6efd; color: white; border-radius: 15px; font-weight: bold; }
        """)

        central = QWidget()
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(30, 20, 30, 30)

        # Header
        header_layout = QHBoxLayout()
        title = QLabel("Tableau de Bord DocuGen")
        title.setFont(QFont("Arial", 22, QFont.Bold))
        btn_help = QPushButton("?")
        btn_help.setObjectName("helpBtn"); btn_help.setFixedSize(30, 30)
        btn_help.clicked.connect(self.show_help)
        header_layout.addWidget(title)
        header_layout.addStretch()
        header_layout.addWidget(btn_help)
        main_layout.addLayout(header_layout)

        # Étape 1 : Modèles récents
        self.card_tpl = QFrame(); self.card_tpl.setObjectName("card")
        l1 = QVBoxLayout(self.card_tpl)
        l1.addWidget(QLabel("<b>1. SÉLECTION DU MODÈLE WORD</b>"))
        
        h_tpl = QHBoxLayout()
        self.combo_tpl = QComboBox()
        self.combo_tpl.setPlaceholderText("Historique des modèles utilisés...")
        self.combo_tpl.currentIndexChanged.connect(self.on_combo_change)
        h_tpl.addWidget(self.combo_tpl, 4)
        
        btn_browse_t = QPushButton("Parcourir...")
        btn_browse_t.clicked.connect(self.select_template)
        h_tpl.addWidget(btn_browse_t, 1)
        l1.addLayout(h_tpl)
        
        self.lbl_tags = QLabel("Balises détectées : -")
        self.lbl_tags.setStyleSheet("color: #0d6efd; font-size: 11px;")
        l1.addWidget(self.lbl_tags)
        main_layout.addWidget(self.card_tpl)

        # Étape 2 : Données
        self.card_ex = QFrame(); self.card_ex.setObjectName("card")
        l2 = QVBoxLayout(self.card_ex)
        l2.addWidget(QLabel("<b>2. SOURCE DE DONNÉES & APERÇU</b>"))
        btn_e = QPushButton("Charger un nouveau fichier Excel")
        btn_e.clicked.connect(self.select_excel)
        l2.addWidget(btn_e)
        self.table = QTableWidget()
        self.table.setMaximumHeight(200)
        l2.addWidget(self.table)
        main_layout.addWidget(self.card_ex)

        main_layout.addStretch()
        
        # Zone d'Action et Résultats
        self.progress = QProgressBar()
        self.progress.hide()
        main_layout.addWidget(self.progress)
        
        self.layout_actions = QHBoxLayout()
        self.btn_run = QPushButton("GÉNÉRER LE DOCUMENT")
        self.btn_run.setObjectName("primaryBtn"); self.btn_run.setFixedHeight(50)
        self.btn_run.setEnabled(False)
        self.btn_run.clicked.connect(self.start)
        
        self.btn_open_result = QPushButton("OUVRIR LE FICHIER GÉNÉRÉ")
        self.btn_open_result.setObjectName("openBtn"); self.btn_open_result.setFixedHeight(50)
        self.btn_open_result.hide() # Caché par défaut
        self.btn_open_result.clicked.connect(self.open_last_file)
        
        self.layout_actions.addWidget(self.btn_run, 2)
        self.layout_actions.addWidget(self.btn_open_result, 1)
        main_layout.addLayout(self.layout_actions)

        self.setCentralWidget(central)

    # --- GESTION DE L'HISTORIQUE ---
    def save_session(self):
        # On récupère l'historique existant pour ne pas le perdre
        history = [self.combo_tpl.itemText(i) for i in range(self.combo_tpl.count())]
        if self.template_path and self.template_path not in history:
            history.insert(0, self.template_path)
            
        config = {
            "template_history": history[:5], # Garde les 5 derniers
            "last_excel": self.excel_path
        }
        with open(CONFIG_FILE, "w") as f:
            json.dump(config, f)

    def load_session(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    cfg = json.load(f)
                    history = cfg.get("template_history", [])
                    self.combo_tpl.addItems(history)
                    if history:
                        self.set_template(history[0])
                    
                    ex = cfg.get("last_excel", "")
                    if os.path.exists(ex):
                        self.set_excel(ex)
            except: pass

    def on_combo_change(self):
        path = self.combo_tpl.currentText()
        if os.path.exists(path):
            self.set_template(path)

    # --- ACTIONS ---
    def set_template(self, path):
        self.template_path = path
        tags = MailGenerator.extract_tags(path)
        self.lbl_tags.setText(f"Colonnes requises : {', '.join(tags)}")
        self.check_ready()

    def select_template(self):
        p, _ = QFileDialog.getOpenFileName(self, "Word", "", "Word (*.docx)")
        if p:
            # Ajouter au combo si nouveau
            if self.combo_tpl.findText(p) == -1:
                self.combo_tpl.insertItem(0, p)
            self.combo_tpl.setCurrentIndex(0)
            self.set_template(p)
            self.save_session()

    def select_excel(self):
        p, _ = QFileDialog.getOpenFileName(self, "Excel", "", "Excel (*.xlsx)")
        if p: self.set_excel(p)

    def set_excel(self, path):
        try:
            df = pd.read_excel(path)
            self.excel_path = path
            self.table.setColumnCount(len(df.columns))
            self.table.setRowCount(min(5, len(df)))
            self.table.setHorizontalHeaderLabels(df.columns)
            for i in range(min(5, len(df))):
                for j in range(len(df.columns)):
                    self.table.setItem(i, j, QTableWidgetItem(str(df.iloc[i, j])))
            self.save_session()
            self.check_ready()
        except: pass

    def check_ready(self):
        self.btn_run.setEnabled(bool(self.template_path and self.excel_path))

    def show_help(self):
        help_dialog = QDialog(self)
        help_dialog.setWindowTitle("Aide DocuGen Pro")
        l = QVBoxLayout(help_dialog)
        tb = QTextBrowser()
        tb.setHtml("""
            <h3>📖 Aide rapide</h3>
            <p>1. Préparez un Word avec des balises comme <b>{{Entreprise}}</b>.</p>
            <p>2. Utilisez les balises auto : <b>{{civ_titre}}</b> (ex: Madame la Directrice).</p>
            <p>3. Dans l'Excel, prévoyez une colonne <b>Sexe</b> (H/F) et <b>Fonction</b>.</p>
            <p>4. Lancez la génération pour obtenir un fichier prêt à imprimer.</p>
        """)
        l.addWidget(tb); help_dialog.resize(400, 300); help_dialog.exec()
    
    def start(self):
        out = QFileDialog.getExistingDirectory(self, "Dossier de sortie")
        if not out: return
        self.progress.show()
        self.btn_run.setEnabled(False)
        try:
            gen = MailGenerator(self.template_path, self.excel_path, out)
            self.last_generated_path = gen.run(lambda v, t: (self.progress.setMaximum(t), self.progress.setValue(v)))
            
            self.btn_open_result.show() # Afficher le bouton ouvrir
            QMessageBox.information(self, "Succès", "Document généré avec succès !")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", str(e))
        finally:
            self.progress.hide()
            self.btn_run.setEnabled(True)

    def open_last_file(self):
        """Ouvre le fichier avec l'application par défaut (Word)."""
        if self.last_generated_path and os.path.exists(self.last_generated_path):
            if os.name == 'nt': # Windows
                os.startfile(self.last_generated_path)
            elif os.name == 'posix': # Mac/Linux
                subprocess.call(['open', self.last_generated_path])