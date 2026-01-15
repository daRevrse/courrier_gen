import os
import pandas as pd
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLabel, QFileDialog, QProgressBar,
                             QMessageBox, QFrame, QTableWidget, QTableWidgetItem,
                             QHeaderView, QStackedWidget, QTextEdit, QSplitter,
                             QCheckBox)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QTextCharFormat, QColor
from logic.processor import MailGenerator
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH


class WordPreview(QTextEdit):
    """Prévisualisation du document Word (lecture seule)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.doc_obj = None
        self.setStyleSheet("""
            QTextEdit {
                border: 1px solid #e2e8f0;
                border-radius: 10px;
                background: white;
                padding: 20px;
                font-family: 'Calibri', 'Arial', sans-serif;
                font-size: 11pt;
            }
        """)

    def load_from_docx(self, doc_path):
        """Charge et affiche le contenu du document Word avec les styles."""
        self.doc_obj = Document(doc_path)
        self.clear()
        cursor = self.textCursor()

        for para in self.doc_obj.paragraphs:
            # Gérer l'alignement
            block_format = cursor.blockFormat()
            if para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                block_format.setAlignment(Qt.AlignCenter)
            elif para.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                block_format.setAlignment(Qt.AlignRight)
            else:
                block_format.setAlignment(Qt.AlignLeft)
            cursor.setBlockFormat(block_format)

            # Appliquer les styles des runs
            for run in para.runs:
                char_format = QTextCharFormat()

                if run.font.name:
                    char_format.setFontFamily(run.font.name)
                if run.font.size:
                    char_format.setFontPointSize(run.font.size.pt)
                if run.bold:
                    char_format.setFontWeight(QFont.Bold)
                if run.italic:
                    char_format.setFontItalic(True)
                if run.underline:
                    char_format.setFontUnderline(True)
                if run.font.color and run.font.color.rgb:
                    rgb = run.font.color.rgb
                    color = QColor(rgb[0], rgb[1], rgb[2])
                    char_format.setForeground(color)

                # Surligner les balises
                if '{{' in run.text and '}}' in run.text:
                    char_format.setBackground(QColor(254, 243, 199))

                cursor.insertText(run.text, char_format)

            cursor.insertBlock()

    def get_document(self):
        """Retourne l'objet Document."""
        return self.doc_obj


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DocuGen Pro v2.5")
        self.setMinimumSize(1200, 800)

        self.template_path = None
        self.current_df = None
        self.last_generated_path = None
        self.template_doc = None

        self.setup_ui()
        self.apply_styles()

    def apply_styles(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #f8fafc; }
            #SideBar { background-color: #0f172a; min-width: 250px; border-right: 1px solid #e2e8f0; }
            #NavBtn {
                background: transparent; color: #cbd5e1; text-align: left;
                padding: 15px 25px; border: none; font-size: 14px; border-radius: 8px; margin: 5px 10px;
            }
            #NavBtn[active="true"] { background: #3b82f6; color: white; font-weight: bold; }
            #NavBtn:hover:not([active="true"]) { background: #1e293b; }

            #MainCard { background: white; border-radius: 15px; border: 1px solid #e2e8f0; }
            #HeaderTitle { font-size: 24px; color: #1e293b; font-weight: bold; }

            QTableWidget { border: 1px solid #e2e8f0; border-radius: 10px; background: white; }
            QHeaderView::section { background-color: #f1f5f9; padding: 8px; border: none; font-weight: bold; }

            QProgressBar { border: none; background: #f1f5f9; border-radius: 6px; height: 12px; }
            QProgressBar::chunk { background: #10b981; border-radius: 6px; }

            QCheckBox {
                spacing: 8px;
                color: #1e293b;
                font-size: 13px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 4px;
                border: 2px solid #cbd5e1;
                background: white;
            }
            QCheckBox::indicator:checked {
                background: #3b82f6;
                border-color: #3b82f6;
            }
        """)

    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)

        # SIDEBAR
        sidebar = QFrame()
        sidebar.setObjectName("SideBar")
        s_layout = QVBoxLayout(sidebar)

        logo = QLabel("DOCUGEN PRO")
        logo.setStyleSheet("color: white; font-size: 20px; font-weight: 900; padding: 25px 20px;")
        s_layout.addWidget(logo)

        self.btn_s1 = QPushButton("  1. Configuration Modele")
        self.btn_s1.setObjectName("NavBtn")
        self.btn_s1.setProperty("active", True)
        self.btn_s2 = QPushButton("  2. Liste des Destinataires")
        self.btn_s2.setObjectName("NavBtn")
        self.btn_s3 = QPushButton("  3. Generation Finale")
        self.btn_s3.setObjectName("NavBtn")

        for b in [self.btn_s1, self.btn_s2, self.btn_s3]:
            s_layout.addWidget(b)
        s_layout.addStretch()
        layout.addWidget(sidebar)

        # CONTENU
        content_pane = QWidget()
        c_layout = QVBoxLayout(content_pane)
        c_layout.setContentsMargins(40, 40, 40, 40)

        self.lbl_title = QLabel("Chargement du Modele")
        self.lbl_title.setObjectName("HeaderTitle")
        c_layout.addWidget(self.lbl_title)
        c_layout.addSpacing(20)

        self.pages = QStackedWidget()
        self.pages.addWidget(self.page_word())
        self.pages.addWidget(self.page_excel())
        self.pages.addWidget(self.page_gen())
        c_layout.addWidget(self.pages)

        # Navigation
        nav = QHBoxLayout()
        self.btn_prev = QPushButton("Retour")
        self.btn_prev.setStyleSheet("""
            QPushButton {
                background: white;
                color: #64748b;
                border: 2px solid #e2e8f0;
                border-radius: 8px;
                padding: 10px 24px;
                font-weight: 600;
                font-size: 13px;
            }
            QPushButton:hover {
                background: #f8fafc;
                border-color: #cbd5e1;
                color: #475569;
            }
        """)
        self.btn_prev.clicked.connect(self.go_back)
        self.btn_prev.hide()

        self.btn_next = QPushButton("Suivant")
        self.btn_next.setStyleSheet("""
            QPushButton {
                background: #3b82f6;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px 24px;
                font-weight: 600;
                font-size: 13px;
            }
            QPushButton:hover {
                background: #2563eb;
            }
        """)
        self.btn_next.clicked.connect(self.go_next)

        nav.addWidget(self.btn_prev)
        nav.addStretch()
        nav.addWidget(self.btn_next)
        c_layout.addLayout(nav)

        layout.addWidget(content_pane, 1)

    def page_word(self):
        f = QFrame()
        f.setObjectName("MainCard")
        l = QVBoxLayout(f)
        l.setContentsMargins(25, 25, 25, 25)

        # Bouton de chargement
        btn_load = QPushButton("Selectionner le modele Word (.docx)")
        btn_load.setStyleSheet("""
            QPushButton {
                background: #3b82f6;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 12px 24px;
                font-weight: 600;
                font-size: 14px;
            }
            QPushButton:hover {
                background: #2563eb;
            }
        """)
        btn_load.clicked.connect(self.load_template_action)
        l.addWidget(btn_load)

        l.addSpacing(10)

        # Info label
        self.template_info = QLabel("Aucun modele charge")
        self.template_info.setStyleSheet("color: #64748b; font-size: 12px;")
        l.addWidget(self.template_info)

        l.addSpacing(10)

        # Splitter pour prévisualisation + info
        splitter = QSplitter(Qt.Horizontal)

        # Prévisualisation (lecture seule)
        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)
        preview_layout.setContentsMargins(0, 0, 5, 0)

        preview_label = QLabel("Previsualisation du modele (lecture seule)")
        preview_label.setStyleSheet("font-weight: bold; color: #1e293b; margin-bottom: 5px;")
        preview_layout.addWidget(preview_label)

        self.word_preview = WordPreview()
        preview_layout.addWidget(self.word_preview)

        # Panneau d'informations
        info_widget = QWidget()
        info_layout = QVBoxLayout(info_widget)
        info_layout.setContentsMargins(5, 0, 0, 0)

        info_label = QLabel("Informations")
        info_label.setStyleSheet("font-weight: bold; color: #1e293b; margin-bottom: 5px;")
        info_layout.addWidget(info_label)

        self.doc_metadata = QLabel("Chargez un document pour voir les informations")
        self.doc_metadata.setWordWrap(True)
        self.doc_metadata.setAlignment(Qt.AlignTop)
        self.doc_metadata.setStyleSheet("""
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 15px;
            color: #64748b;
            font-size: 12px;
        """)
        info_layout.addWidget(self.doc_metadata)

        splitter.addWidget(preview_widget)
        splitter.addWidget(info_widget)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)

        l.addWidget(splitter)

        return f

    def page_excel(self):
        f = QFrame()
        f.setObjectName("MainCard")
        l = QVBoxLayout(f)
        l.setContentsMargins(25, 25, 25, 25)

        # Header
        header_layout = QHBoxLayout()
        btn = QPushButton("Importer la base Excel")
        btn.setStyleSheet("""
            QPushButton {
                background: #3b82f6;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 12px 24px;
                font-weight: 600;
                font-size: 14px;
            }
            QPushButton:hover {
                background: #2563eb;
            }
        """)
        btn.clicked.connect(self.load_excel_action)
        header_layout.addWidget(btn)

        self.excel_info = QLabel("Aucun fichier Excel charge")
        self.excel_info.setStyleSheet("color: #64748b; font-size: 12px; margin-left: 15px;")
        header_layout.addWidget(self.excel_info)
        header_layout.addStretch()

        l.addLayout(header_layout)
        l.addSpacing(15)

        # Table
        self.data_table = QTableWidget()
        self.data_table.setAlternatingRowColors(True)
        self.data_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #e2e8f0;
                border-radius: 10px;
                background: white;
                gridline-color: #f1f5f9;
            }
            QTableWidget::item { padding: 8px; }
            QTableWidget::item:selected {
                background: #dbeafe;
                color: #1e40af;
            }
            QHeaderView::section {
                background-color: #f8fafc;
                padding: 10px;
                border: none;
                border-bottom: 2px solid #e2e8f0;
                font-weight: bold;
                color: #1e293b;
            }
        """)
        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.data_table.setEditTriggers(QTableWidget.DoubleClicked)
        l.addWidget(self.data_table)

        return f

    def page_gen(self):
        f = QFrame()
        f.setObjectName("MainCard")
        l = QVBoxLayout(f)
        l.setAlignment(Qt.AlignCenter)
        l.setSpacing(25)

        title = QLabel("PRET POUR LA GENERATION")
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #1e293b;")
        title.setAlignment(Qt.AlignCenter)
        l.addWidget(title)

        self.summary = QLabel("Le systeme va fusionner votre modele Word avec les donnees Excel.\nTous les styles et la mise en forme seront preserves.")
        self.summary.setStyleSheet("color: #64748b; font-size: 14px; text-align: center;")
        self.summary.setAlignment(Qt.AlignCenter)
        self.summary.setWordWrap(True)
        l.addWidget(self.summary)

        # Option de saut de page
        page_break_container = QWidget()
        page_break_layout = QVBoxLayout(page_break_container)
        page_break_layout.setSpacing(10)

        self.page_break_checkbox = QCheckBox("Ajouter un saut de page entre chaque courrier")
        self.page_break_checkbox.setChecked(True)
        self.page_break_checkbox.setStyleSheet("font-weight: 600;")
        page_break_layout.addWidget(self.page_break_checkbox)

        page_break_hint = QLabel("Decochez si votre modele gere deja les sauts de page")
        page_break_hint.setStyleSheet("color: #94a3b8; font-size: 11px; margin-left: 26px;")
        page_break_layout.addWidget(page_break_hint)

        l.addWidget(page_break_container)

        l.addSpacing(20)

        # Barre de progression
        self.progress = QProgressBar()
        self.progress.setFixedHeight(20)
        self.progress.setStyleSheet("""
            QProgressBar {
                border: 2px solid #e2e8f0;
                border-radius: 10px;
                background: #f8fafc;
                text-align: center;
                font-weight: bold;
                color: #1e293b;
            }
            QProgressBar::chunk {
                background: #3b82f6;
                border-radius: 8px;
            }
        """)
        self.progress.hide()
        l.addWidget(self.progress)

        # Bouton de génération
        self.btn_run = QPushButton("GENERER LES COURRIERS")
        self.btn_run.setFixedHeight(54)
        self.btn_run.setStyleSheet("""
            QPushButton {
                background: #3b82f6;
                color: white;
                border: none;
                border-radius: 12px;
                padding: 16px 32px;
                font-weight: bold;
                font-size: 15px;
                letter-spacing: 0.5px;
            }
            QPushButton:hover {
                background: #2563eb;
            }
            QPushButton:disabled {
                background: #cbd5e1;
                color: #94a3b8;
            }
        """)
        self.btn_run.clicked.connect(self.execute_generation)
        l.addWidget(self.btn_run)

        # Bouton d'ouverture
        self.btn_open = QPushButton("OUVRIR LE DOCUMENT")
        self.btn_open.setFixedHeight(48)
        self.btn_open.setStyleSheet("""
            QPushButton {
                background: #10b981;
                color: white;
                border: none;
                border-radius: 10px;
                padding: 14px 28px;
                font-weight: 600;
                font-size: 14px;
            }
            QPushButton:hover {
                background: #059669;
            }
        """)
        self.btn_open.hide()
        self.btn_open.clicked.connect(self.open_result)
        l.addWidget(self.btn_open)

        return f

    # LOGIQUE
    def update_state(self):
        idx = self.pages.currentIndex()
        titles = ["Configuration Modele", "Destinataires", "Generation"]
        self.lbl_title.setText(titles[idx])

        btns = [self.btn_s1, self.btn_s2, self.btn_s3]
        for i, b in enumerate(btns):
            b.setProperty("active", i == idx)
            b.style().unpolish(b)
            b.style().polish(b)

        self.btn_prev.setVisible(idx > 0)
        self.btn_next.setVisible(idx < 2)

    def go_next(self):
        if self.pages.currentIndex() == 0 and not self.template_path:
            return
        if self.pages.currentIndex() == 1:
            self.sync_table_to_df()
        self.pages.setCurrentIndex(self.pages.currentIndex() + 1)
        self.update_state()

    def go_back(self):
        self.pages.setCurrentIndex(self.pages.currentIndex() - 1)
        self.update_state()

    def load_template_action(self):
        path, _ = QFileDialog.getOpenFileName(self, "Selectionner un modele Word", "", "Word (*.docx)")
        if path:
            self.template_path = path
            self.word_preview.load_from_docx(path)
            self.template_doc = self.word_preview.get_document()

            # Mettre à jour les informations
            doc = self.template_doc
            num_sections = len(doc.sections)
            num_paragraphs = len(doc.paragraphs)
            num_tables = len(doc.tables)

            # Extraire les balises
            import re
            tags = set()
            for para in doc.paragraphs:
                tags.update(re.findall(r'\{\{\s*(\w+)\s*\}\}', para.text))
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        tags.update(re.findall(r'\{\{\s*(\w+)\s*\}\}', cell.text))

            info_text = f"""<b>Fichier:</b><br/>
            <span style='font-size: 10px;'>{path}</span><br/><br/>
            <b>Sections:</b> {num_sections}<br/>
            <b>Paragraphes:</b> {num_paragraphs}<br/>
            <b>Tableaux:</b> {num_tables}<br/><br/>
            <b>Balises detectees:</b> {len(tags)}<br/>
            """

            if tags:
                info_text += "<div style='margin-top: 10px;'>"
                for tag in sorted(tags):
                    info_text += f"<span style='background: #dbeafe; color: #1e40af; padding: 2px 6px; border-radius: 4px; margin: 2px; display: inline-block; font-size: 10px;'>{{{{ {tag} }}}}</span> "
                info_text += "</div>"

            self.doc_metadata.setText(info_text)
            self.template_info.setText(f"Modele charge: {path.split('/')[-1]}")
            self.template_info.setStyleSheet("color: #10b981; font-size: 12px; font-weight: bold;")

    def load_excel_action(self):
        path, _ = QFileDialog.getOpenFileName(self, "Selectionner un fichier Excel", "", "Excel (*.xlsx)")
        if path:
            self.current_df = pd.read_excel(path)
            self.excel_info.setText(f"{len(self.current_df)} lignes | {len(self.current_df.columns)} colonnes | {path.split('/')[-1]}")
            self.excel_info.setStyleSheet("color: #10b981; font-size: 12px; margin-left: 15px; font-weight: bold;")
            self.refresh_table()

    def refresh_table(self):
        df = self.current_df
        self.data_table.setColumnCount(len(df.columns))
        self.data_table.setRowCount(len(df))
        self.data_table.setHorizontalHeaderLabels(df.columns)
        for i in range(len(df)):
            for j in range(len(df.columns)):
                self.data_table.setItem(i, j, QTableWidgetItem(str(df.iloc[i, j])))

    def sync_table_to_df(self):
        data = []
        for i in range(self.data_table.rowCount()):
            row = [self.data_table.item(i, j).text() if self.data_table.item(i, j) else "" for j in range(self.data_table.columnCount())]
            data.append(row)
        self.current_df = pd.DataFrame(data, columns=[self.data_table.horizontalHeaderItem(i).text() for i in range(self.data_table.columnCount())])

    def execute_generation(self):
        dest = QFileDialog.getExistingDirectory(self, "Dossier de sortie")
        if not dest:
            return

        self.btn_run.setEnabled(False)
        self.btn_run.setText("Generation en cours...")
        self.progress.show()
        self.progress.setFormat("%v / %m courriers generes (%p%)")

        try:
            gen = MailGenerator(self.template_path, None, dest)

            # Utiliser la checkbox pour contrôler les sauts de page
            add_page_breaks = self.page_break_checkbox.isChecked()

            self.last_generated_path = gen.run_direct(
                self.current_df,
                add_page_breaks,
                lambda v, t: (self.progress.setMaximum(t), self.progress.setValue(v))
            )

            # Afficher le bouton de succès
            self.btn_run.hide()
            self.btn_open.show()

            # Message de succès
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Generation reussie")
            msg.setText(f"<h3 style='color: #10b981;'>Generation terminee avec succes !</h3>")
            msg.setInformativeText(f"<p>{len(self.current_df)} courriers ont ete generes.</p><p><b>Fichier:</b><br/>{self.last_generated_path}</p>")
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec()

        except Exception as e:
            self.btn_run.setEnabled(True)
            self.btn_run.setText("GENERER LES COURRIERS")
            QMessageBox.critical(self, "Erreur", f"Une erreur est survenue:\n\n{str(e)}")

    def open_result(self):
        if self.last_generated_path:
            os.startfile(self.last_generated_path)
