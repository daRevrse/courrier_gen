import os
import pandas as pd
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLabel, QFileDialog, QProgressBar,
                             QMessageBox, QFrame, QTableWidget, QTableWidgetItem,
                             QHeaderView, QStackedWidget, QTextEdit, QCheckBox,
                             QToolBar, QFontComboBox, QSpinBox)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QFont, QTextCharFormat, QColor, QAction, QTextCursor, QTextBlockFormat
from logic.processor import MailGenerator
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH


class RichTextEditor(QTextEdit):
    """Éditeur de texte riche avec barre d'outils."""

    def __init__(self, parent=None):
        super().__init__(parent)
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
        try:
            self.doc_obj = Document(doc_path)
            self.original_path = doc_path

            # Bloquer les signaux et désactiver les mises à jour pendant le chargement
            self.blockSignals(True)
            self.setUpdatesEnabled(False)

            # Vider le contenu
            self.clear()

            # Obtenir le curseur
            cursor = self.textCursor()
            cursor.beginEditBlock()  # Grouper toutes les modifications en une seule opération

            try:
                for para in self.doc_obj.paragraphs:
                    # Gérer l'alignement du paragraphe
                    block_format = QTextBlockFormat()
                    try:
                        if para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                            block_format.setAlignment(Qt.AlignCenter)
                        elif para.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                            block_format.setAlignment(Qt.AlignRight)
                        elif para.alignment == WD_ALIGN_PARAGRAPH.LEFT:
                            block_format.setAlignment(Qt.AlignLeft)
                        cursor.setBlockFormat(block_format)
                    except:
                        pass

                    # Traiter chaque run
                    if para.runs:
                        for run in para.runs:
                            text = run.text if run.text else ""
                            if not text:
                                continue

                            # Créer le format de caractère
                            char_format = QTextCharFormat()

                            # Police
                            try:
                                if run.font.name:
                                    char_format.setFontFamily(run.font.name)
                            except:
                                pass

                            # Taille
                            try:
                                if run.font.size:
                                    char_format.setFontPointSize(run.font.size.pt)
                            except:
                                pass

                            # Gras
                            try:
                                if run.bold:
                                    char_format.setFontWeight(QFont.Weight.Bold)
                                else:
                                    char_format.setFontWeight(QFont.Weight.Normal)
                            except:
                                pass

                            # Italique
                            try:
                                if run.italic:
                                    char_format.setFontItalic(True)
                            except:
                                pass

                            # Souligné
                            try:
                                if run.underline:
                                    char_format.setFontUnderline(True)
                            except:
                                pass

                            # Couleur du texte
                            try:
                                if run.font.color and run.font.color.rgb:
                                    rgb = run.font.color.rgb
                                    if rgb:
                                        color = QColor(rgb >> 16 & 0xFF, rgb >> 8 & 0xFF, rgb & 0xFF)
                                        char_format.setForeground(color)
                            except:
                                pass

                            # Surligner les balises
                            if '{{' in text and '}}' in text:
                                char_format.setBackground(QColor(254, 243, 199))

                            # Insérer le texte avec le format
                            cursor.insertText(text, char_format)

                    # Nouvelle ligne pour le prochain paragraphe
                    cursor.insertBlock()

            finally:
                cursor.endEditBlock()  # Terminer le groupement des modifications

                # Réactiver les mises à jour et les signaux
                self.setUpdatesEnabled(True)
                self.blockSignals(False)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.setUpdatesEnabled(True)
            self.blockSignals(False)
            self.setPlainText(f"Erreur lors du chargement du document:\n{str(e)}")

    def get_document(self):
        """Retourne l'objet Document."""
        return self.doc_obj

    def get_original_path(self):
        """Retourne le chemin du fichier original."""
        return self.original_path


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DocuGen Pro v2.5")
        self.setMinimumSize(1400, 900)

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

            QToolBar {
                background: #f8fafc;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                padding: 4px;
                spacing: 4px;
            }
            QToolButton {
                background: white;
                border: 1px solid #e2e8f0;
                border-radius: 6px;
                padding: 6px;
                margin: 2px;
            }
            QToolButton:hover {
                background: #f1f5f9;
                border-color: #cbd5e1;
            }
            QToolButton:checked {
                background: #dbeafe;
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
        l.setContentsMargins(20, 20, 20, 20)

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

        l.addSpacing(15)

        # Barre d'outils de formatage
        toolbar_label = QLabel("Outils de formatage:")
        toolbar_label.setStyleSheet("font-weight: 600; color: #1e293b; font-size: 12px; margin-bottom: 5px;")
        l.addWidget(toolbar_label)

        toolbar = QToolBar()
        toolbar.setIconSize(QSize(20, 20))
        toolbar.setMovable(False)
        toolbar.setStyleSheet("""
            QToolBar {
                background: #f8fafc;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                padding: 6px;
                spacing: 3px;
            }
            QToolButton {
                background: white;
                border: 1px solid #e2e8f0;
                border-radius: 6px;
                padding: 8px;
                margin: 2px;
                min-width: 32px;
                min-height: 32px;
            }
            QToolButton:hover {
                background: #f1f5f9;
                border-color: #cbd5e1;
            }
            QToolButton:checked {
                background: #dbeafe;
                border-color: #3b82f6;
                color: #3b82f6;
            }
            QFontComboBox, QSpinBox {
                background: white;
                border: 1px solid #e2e8f0;
                border-radius: 6px;
                padding: 6px;
                min-height: 28px;
            }
        """)

        # Police
        self.font_combo = QFontComboBox()
        self.font_combo.setFixedWidth(150)
        self.font_combo.currentFontChanged.connect(self.change_font)
        toolbar.addWidget(self.font_combo)

        # Taille
        self.size_spin = QSpinBox()
        self.size_spin.setRange(8, 72)
        self.size_spin.setValue(11)
        self.size_spin.setFixedWidth(60)
        self.size_spin.valueChanged.connect(self.change_size)
        toolbar.addWidget(self.size_spin)

        toolbar.addSeparator()

        # Gras
        self.bold_action = QAction("B", self)
        self.bold_action.setCheckable(True)
        self.bold_action.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        self.bold_action.triggered.connect(self.toggle_bold)
        toolbar.addAction(self.bold_action)

        # Italique
        self.italic_action = QAction("I", self)
        self.italic_action.setCheckable(True)
        italic_font = QFont("Arial", 10)
        italic_font.setItalic(True)
        self.italic_action.setFont(italic_font)
        self.italic_action.triggered.connect(self.toggle_italic)
        toolbar.addAction(self.italic_action)

        # Souligné
        self.underline_action = QAction("U", self)
        self.underline_action.setCheckable(True)
        underline_font = QFont("Arial", 10)
        underline_font.setUnderline(True)
        self.underline_action.setFont(underline_font)
        self.underline_action.triggered.connect(self.toggle_underline)
        toolbar.addAction(self.underline_action)

        toolbar.addSeparator()

        # Alignement
        align_left = QAction("Gauche", self)
        align_left.triggered.connect(lambda: self.change_alignment(Qt.AlignmentFlag.AlignLeft))
        toolbar.addAction(align_left)

        align_center = QAction("Centre", self)
        align_center.triggered.connect(lambda: self.change_alignment(Qt.AlignmentFlag.AlignCenter))
        toolbar.addAction(align_center)

        align_right = QAction("Droite", self)
        align_right.triggered.connect(lambda: self.change_alignment(Qt.AlignmentFlag.AlignRight))
        toolbar.addAction(align_right)

        l.addWidget(toolbar)

        # Éditeur de texte (agrandi)
        editor_label = QLabel("Apercu du modele avec outils de formatage")
        editor_label.setStyleSheet("font-weight: bold; color: #1e293b; margin-bottom: 5px; margin-top: 10px;")
        l.addWidget(editor_label)

        # Note importante
        note_label = QLabel("Note: L'editeur permet de visualiser et tester les styles. Pour modifier definitivement le modele, editez le fichier Word original.")
        note_label.setStyleSheet("color: #64748b; font-size: 11px; font-style: italic; margin-bottom: 5px;")
        note_label.setWordWrap(True)
        l.addWidget(note_label)

        self.word_editor = RichTextEditor()
        # Connecter le signal après un court délai pour éviter les crashes au chargement
        # self.word_editor.cursorPositionChanged.connect(self.update_format_toolbar)
        l.addWidget(self.word_editor)

        # Métadonnées en bas (plus compact)
        meta_label = QLabel("Informations du document")
        meta_label.setStyleSheet("font-weight: 600; color: #1e293b; font-size: 12px; margin-top: 10px;")
        l.addWidget(meta_label)

        self.doc_metadata = QLabel("Chargez un document pour voir les informations")
        self.doc_metadata.setWordWrap(True)
        self.doc_metadata.setStyleSheet("""
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 10px;
            color: #64748b;
            font-size: 11px;
        """)
        self.doc_metadata.setMaximumHeight(100)
        l.addWidget(self.doc_metadata)

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

        # Bouton pour nouvelle génération
        self.btn_new_gen = QPushButton("NOUVELLE GENERATION")
        self.btn_new_gen.setFixedHeight(48)
        self.btn_new_gen.setStyleSheet("""
            QPushButton {
                background: #3b82f6;
                color: white;
                border: none;
                border-radius: 10px;
                padding: 14px 28px;
                font-weight: 600;
                font-size: 14px;
            }
            QPushButton:hover {
                background: #2563eb;
            }
        """)
        self.btn_new_gen.hide()
        self.btn_new_gen.clicked.connect(self.reset_generation)
        l.addWidget(self.btn_new_gen)

        return f

    # OUTILS DE FORMATAGE
    def update_format_toolbar(self):
        """Met à jour l'état de la toolbar selon le formatage actuel."""
        try:
            cursor = self.word_editor.textCursor()
            char_format = cursor.charFormat()

            # Bloquer les signaux pour éviter les boucles
            self.font_combo.blockSignals(True)
            self.size_spin.blockSignals(True)
            self.bold_action.blockSignals(True)
            self.italic_action.blockSignals(True)
            self.underline_action.blockSignals(True)

            # Mettre à jour les boutons
            self.bold_action.setChecked(char_format.fontWeight() == QFont.Weight.Bold)
            self.italic_action.setChecked(char_format.fontItalic())
            self.underline_action.setChecked(char_format.fontUnderline())

            # Mettre à jour la police et taille
            if char_format.fontFamily():
                self.font_combo.setCurrentFont(QFont(char_format.fontFamily()))
            if char_format.fontPointSize() > 0:
                self.size_spin.setValue(int(char_format.fontPointSize()))

            # Débloquer les signaux
            self.font_combo.blockSignals(False)
            self.size_spin.blockSignals(False)
            self.bold_action.blockSignals(False)
            self.italic_action.blockSignals(False)
            self.underline_action.blockSignals(False)
        except:
            pass

    def change_font(self, font):
        """Change la police du texte sélectionné."""
        try:
            char_format = QTextCharFormat()
            char_format.setFontFamily(font.family())
            self.merge_format(char_format)
        except:
            pass

    def change_size(self, size):
        """Change la taille du texte sélectionné."""
        try:
            char_format = QTextCharFormat()
            char_format.setFontPointSize(size)
            self.merge_format(char_format)
        except:
            pass

    def toggle_bold(self):
        """Active/désactive le gras."""
        try:
            char_format = QTextCharFormat()
            if self.bold_action.isChecked():
                char_format.setFontWeight(QFont.Weight.Bold)
            else:
                char_format.setFontWeight(QFont.Weight.Normal)
            self.merge_format(char_format)
        except:
            pass

    def toggle_italic(self):
        """Active/désactive l'italique."""
        try:
            char_format = QTextCharFormat()
            char_format.setFontItalic(self.italic_action.isChecked())
            self.merge_format(char_format)
        except:
            pass

    def toggle_underline(self):
        """Active/désactive le souligné."""
        try:
            char_format = QTextCharFormat()
            char_format.setFontUnderline(self.underline_action.isChecked())
            self.merge_format(char_format)
        except:
            pass

    def change_alignment(self, alignment):
        """Change l'alignement du paragraphe."""
        try:
            self.word_editor.setAlignment(alignment)
        except:
            pass

    def merge_format(self, char_format):
        """Applique le formatage au texte sélectionné."""
        try:
            cursor = self.word_editor.textCursor()
            if not cursor.hasSelection():
                cursor.select(QTextCursor.WordUnderCursor)
            cursor.mergeCharFormat(char_format)
            self.word_editor.mergeCurrentCharFormat(char_format)
        except:
            pass

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
            try:
                self.template_path = path
                self.word_editor.load_from_docx(path)
                self.template_doc = self.word_editor.get_document()

                # Vérifier que le document a bien été chargé
                if not self.template_doc:
                    raise Exception("Le document n'a pas pu être chargé correctement")

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

                info_text = f"<b>Fichier:</b> {os.path.basename(path)}<br/>"
                info_text += f"<b>Sections:</b> {num_sections} | <b>Paragraphes:</b> {num_paragraphs} | <b>Tableaux:</b> {num_tables}<br/>"
                info_text += f"<b>Balises:</b> {len(tags)}"

                if tags and len(tags) <= 5:
                    info_text += " - " + ", ".join([f"{{{{ {tag} }}}}" for tag in sorted(tags)])

                self.doc_metadata.setText(info_text)
                self.template_info.setText(f"Modele charge: {os.path.basename(path)}")
                self.template_info.setStyleSheet("color: #10b981; font-size: 12px; font-weight: bold;")

            except Exception as e:
                QMessageBox.critical(self, "Erreur de chargement",
                                   f"Impossible de charger le modele Word:\n\n{str(e)}\n\nVerifiez que le fichier est un document Word valide (.docx).")
                self.template_path = None
                self.template_doc = None

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
            # Utiliser le fichier Word original (pas les modifications de l'éditeur)
            gen = MailGenerator(self.template_path, None, dest)

            # Utiliser la checkbox pour contrôler les sauts de page
            add_page_breaks = self.page_break_checkbox.isChecked()

            self.last_generated_path = gen.run_direct(
                self.current_df,
                add_page_breaks,
                lambda v, t: (self.progress.setMaximum(t), self.progress.setValue(v))
            )

            # Afficher les boutons de succès
            self.btn_run.hide()
            self.btn_open.show()
            self.btn_new_gen.show()

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

    def reset_generation(self):
        """Réinitialise l'interface pour permettre une nouvelle génération."""
        self.btn_open.hide()
        self.btn_new_gen.hide()
        self.btn_run.show()
        self.btn_run.setEnabled(True)
        self.btn_run.setText("GENERER LES COURRIERS")
        self.progress.hide()
        self.progress.setValue(0)

    def open_result(self):
        if self.last_generated_path:
            os.startfile(self.last_generated_path)
