import sys
import os
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QTabWidget, QLabel, QPushButton, QFrame, QSplitter, 
    QStatusBar, QMessageBox, QToolBar, QTableWidget, QTableWidgetItem,
    QHeaderView, QLineEdit, QDialog, QFormLayout, QComboBox
)
from PySide6.QtCore import Qt, Slot
from PySide6.QtGui import QAction, QIcon, QFont

# Importaciones locales
from license_manager import obtener_estado_acceso_software
from db_connector import SupabaseConnector
# ... (imports anteriores)
from document_manager import DocumentManager  # Importar el nuevo módulo

class GCDRYSApp(QMainWindow):
    def __init__(self, user_data):
        super().__init__()
        self.user_data = user_data
        self.db = SupabaseConnector()
        
        self.setWindowTitle(f"GCDRYS-2 | Usuario: {user_data['username']} ({user_data['rol']})")
        self.setMinimumSize(1060, 730)
        
        self.init_styles()
        self.init_ui()
        self.init_menu_bar()
        self.init_status_bar()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)

        # --- Panel Izquierdo (Sin cambios mayores) ---
        left_panel = QFrame()
        left_panel.setFixedWidth(250) # Un poco más estrecho para dar espacio a la tabla
        left_layout = QVBoxLayout(left_panel)
        left_layout.setAlignment(Qt.AlignTop)

        logo_label = QLabel("GCDRYS-2")
        logo_label.setFont(QFont("Arial", 24, QFont.Bold))
        logo_label.setAlignment(Qt.AlignCenter)
        logo_label.setStyleSheet("color: #2d5bd1; margin-bottom: 10px;")
        
        left_layout.addWidget(logo_label)
        left_layout.addSpacing(20)

        # Info de Usuario
        user_info_frame = QFrame()
        user_info_frame.setStyleSheet("background: #f9f9f9; border: none; padding: 10px;")
        user_layout = QVBoxLayout(user_info_frame)
        user_layout.addWidget(QLabel(f"<b>Usuario:</b> {self.user_data['username']}"))
        user_layout.addWidget(QLabel(f"<b>Rol:</b> {self.user_data['rol']}"))
        left_layout.addWidget(user_info_frame)
        
        left_layout.addStretch()

        btn_admin = QPushButton("Panel de Configuración")
        btn_admin.clicked.connect(self.abrir_panel_configuracion)
        if self.user_data['rol'] in ['admin', 'Superusuario']:
            left_layout.addWidget(btn_admin)

        # --- Panel Derecho: Gestor de Documentos ---
        right_panel = QFrame()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        # Instanciamos el gestor de documentos
        self.doc_manager = DocumentManager(self.db, self.user_data)
        right_layout.addWidget(self.doc_manager)

        main_layout.addWidget(left_panel)
        main_layout.addWidget(right_panel, stretch=1)

    # ... (resto de métodos: init_menu_bar, init_status_bar, abrir_panel_configuracion, etc.)

        # Layout Principal
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(40, 40, 40, 40)
        main_layout.setSpacing(15)

        # --- Logo y Título (Simulado con texto/estilos) ---
        # En PySide6, puedes usar QLabel con QPixmap para poner el logo real de R&S
        # logo_label = QLabel()
        # logo_label.setPixmap(QPixmap("ruta/a/tu/logo.png").scaled(100, 100, Qt.KeepAspectRatio))
        # logo_label.setAlignment(Qt.AlignCenter)
        # main_layout.addWidget(logo_label)
        
        title_label = QLabel("GCDRYS-2")
        title_label.setObjectName("title_label")
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        subtitle_label = QLabel("Gestión y Control Documental")
        subtitle_label.setObjectName("subtitle_label")
        subtitle_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(subtitle_label)

        main_layout.addSpacing(20)

        # --- Campos de Entrada ---
        self.user_input = QLineEdit()
        self.user_input.setPlaceholderText("Usuario")
        
        self.pass_input = QLineEdit()
        self.pass_input.setPlaceholderText("Contraseña")
        self.pass_input.setEchoMode(QLineEdit.Password)

        main_layout.addWidget(self.user_input)
        main_layout.addWidget(self.pass_input)

        # Label de error oculto por defecto
        self.error_label = QLabel("")
        self.error_label.setObjectName("error_label")
        self.error_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.error_label)

        # --- Botón de Login ---
        btn_login = QPushButton("Login")
        btn_login.setObjectName("login_btn")
        btn_login.clicked.connect(self.intentar_login)
        main_layout.addWidget(btn_login)

        main_layout.addStretch()

        # Permitir login con Enter
        self.pass_input.returnPressed.connect(self.intentar_login)
        self.user_input.returnPressed.connect(lambda: self.pass_input.setFocus())

    def intentar_login(self):
        user = self.user_input.text().strip()
        pwd = self.pass_input.text()
        
        if not user or not pwd:
            self.error_label.setText("Por favor, complete todos los campos.")
            return
            
        self.error_label.setText("") # Limpiar errores previos
        self.setCursor(Qt.WaitCursor) # Feedback visual de carga
        
        resultado = self.db.validar_credenciales(user, pwd)
        
        self.setCursor(Qt.ArrowCursor)
        
        if resultado:
            self.user_data = {
                "id": resultado[0],
                "username": resultado[1],
                "rol": resultado[2]
            }
            self.accept()
        else:
            self.error_label.setText("Usuario o contraseña incorrectos.")
            self.pass_input.clear()
            self.pass_input.setFocus()


class GCDRYSApp(QMainWindow):
    def __init__(self, user_data):
        super().__init__()
        self.user_data = user_data
        self.db = SupabaseConnector()
        
        self.setWindowTitle(f"GCDRYS-2 | Usuario: {user_data['username']} ({user_data['rol']})")
        self.setMinimumSize(1060, 730)
        
        self.init_styles()
        self.init_ui()
        self.init_menu_bar()
        self.init_status_bar()

    def init_styles(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #f0f0f0; }
            QFrame { background-color: white; border: 1px solid #dcdcdc; border-radius: 4px; }
            QLabel { color: #333333; }
            QPushButton {
                background-color: #2d5bd1; color: white; border: none;
                padding: 8px 16px; border-radius: 4px; font-weight: bold;
            }
            QPushButton:hover { background-color: #1e45a8; }
            QTabWidget::pane { border: 1px solid #dcdcdc; background: white; }
            QTabBar::tab {
                background: #e0e0e0; padding: 8px 12px; margin-right: 2px;
            }
            QTabBar::tab:selected { background: white; border-bottom: 2px solid #2d5bd1; }
            QTableWidget { gridline-color: #dcdcdc; }
            QTableWidget::item { padding: 5px; }
        """)

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)

        # --- Panel Izquierdo ---
        left_panel = QFrame()
        left_panel.setFixedWidth(320)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setAlignment(Qt.AlignTop)

        logo_label = QLabel("GCDRYS-2")
        logo_label.setFont(QFont("Arial", 26, QFont.Bold))
        logo_label.setAlignment(Qt.AlignCenter)
        logo_label.setStyleSheet("color: #2d5bd1; margin-bottom: 10px;")
        
        info_label = QLabel("Sistema de Gestión Documental")
        info_label.setAlignment(Qt.AlignCenter)
        info_label.setStyleSheet("color: #666; font-size: 12px;")

        left_layout.addWidget(logo_label)
        left_layout.addWidget(info_label)
        left_layout.addStretch()

        # Info de Usuario Logueado
        user_info_frame = QFrame()
        user_info_frame.setStyleSheet("background: #f9f9f9; border: none;")
        user_layout = QVBoxLayout(user_info_frame)
        user_layout.addWidget(QLabel(f"Usuario: <b>{self.user_data['username']}</b>"))
        user_layout.addWidget(QLabel(f"Rol: <b>{self.user_data['rol']}</b>"))
        user_layout.addWidget(QLabel("Estado: <font color='green'>Conectado a Supabase</font>"))
        left_layout.addWidget(user_info_frame)
        left_layout.addStretch()

        btn_admin = QPushButton("Panel de Configuración")
        btn_admin.clicked.connect(self.abrir_panel_configuracion)
        # Solo mostrar si tiene permisos (ejemplo básico)
        if self.user_data['rol'] in ['admin', 'Superusuario']:
            left_layout.addWidget(btn_admin)

        # --- Panel Derecho ---
        right_panel = QFrame()
        right_layout = QVBoxLayout(right_panel)
        
        welcome_label = QLabel(f"Bienvenido, {self.user_data['username']}.")
        welcome_label.setAlignment(Qt.AlignCenter)
        welcome_label.setStyleSheet("font-size: 18px; color: #888;")
        right_layout.addWidget(welcome_label)
        right_layout.addStretch()

        main_layout.addWidget(left_panel)
        main_layout.addWidget(right_panel, stretch=1)

    def init_menu_bar(self):
        menubar = self.menuBar()
        file_menu = menubar.addMenu("Archivo")
        exit_action = QAction("Salir", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

    def init_status_bar(self):
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Listo")

    @Slot()
    def abrir_panel_configuracion(self):
        dialog = ConfigDialog(self, self.db)
        dialog.exec()

class ConfigDialog(QWidget):
    def __init__(self, parent=None, db_connector=None):
        super().__init__(parent)
        self.setWindowTitle("Panel de Configuración")
        self.resize(980, 680)
        self.db = db_connector
        
        layout = QVBoxLayout(self)
        self.tabs = QTabWidget()
        
        self.tabs.addTab(EmpresaTab(), "Empresa")
        self.tabs.addTab(LicenciaTab(), "Licencia")
        self.tabs.addTab(AreasTab(self.db), "Áreas")
        
        layout.addWidget(self.tabs)
        
        btn_close = QPushButton("Cerrar")
        btn_close.setFixedWidth(100)
        btn_close.clicked.connect(self.close)
        layout.addWidget(btn_close, alignment=Qt.AlignRight)

class EmpresaTab(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Configuración de Datos de la Empresa"))

class LicenciaTab(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Gestión de Licencias y Activación"))

class AreasTab(QWidget):
    def __init__(self, db_connector):
        super().__init__()
        self.db = db_connector
        layout = QVBoxLayout(self)
        
        header = QLabel("Administración de Áreas Corporativas")
        header.setFont(QFont("Arial", 12, QFont.Bold))
        layout.addWidget(header)
        
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["ID", "Nombre del Área"])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        layout.addWidget(self.table)
        
        btn_refresh = QPushButton("Actualizar desde Supabase")
        btn_refresh.clicked.connect(self.cargar_areas)
        layout.addWidget(btn_refresh, alignment=Qt.AlignRight)
        
        self.cargar_areas()

    def cargar_areas(self):
        self.table.setRowCount(0)
        areas = self.db.obtener_areas()
        
        for row_num, area in enumerate(areas):
            self.table.insertRow(row_num)
            self.table.setItem(row_num, 0, QTableWidgetItem(str(area['id'])))
            self.table.setItem(row_num, 1, QTableWidgetItem(area['nombre']))

class LoginDialog(QDialog):
    def __init__(self, db_connector, parent=None):
        super().__init__(parent)
        self.db = db_connector
        self.user_data = None
        
        self.setWindowTitle("GCDRYS-2")
        self.setFixedSize(400, 550)
        self.setStyleSheet("""
            QDialog {
                background-color: #ffffff;
            }
            QLabel#title_label {
                color: #2d5bd1;
                font-size: 28px;
                font-weight: bold;
                letter-spacing: 1px;
            }
            QLabel#subtitle_label {
                color: #666666;
                font-size: 14px;
                margin-bottom: 20px;
            }
            QLineEdit {
                border: none;
                border-bottom: 2px solid #e0e0e0;
                padding: 10px 0;
                font-size: 16px;
                color: #333;
                background: transparent;
            }
            QLineEdit:focus {
                border-bottom: 2px solid #2d5bd1;
            }
            QPushButton#login_btn {
                background-color: #2d5bd1;
                color: white;
                border: none;
                padding: 12px;
                border-radius: 6px;
                font-size: 16px;
                font-weight: bold;
                margin-top: 20px;
            }
            QPushButton#login_btn:hover {
                background-color: #1e45a8;
            }
            QLabel#error_label {
                color: #d32f2f;
                font-size: 12px;
                margin-top: 10px;
            }
        """)

        # Layout Principal
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(40, 40, 40, 40)
        main_layout.setSpacing(15)

        # Título
        title_label = QLabel("GCDRYS-2")
        title_label.setObjectName("title_label")
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        subtitle_label = QLabel("Gestión y Control Documental")
        subtitle_label.setObjectName("subtitle_label")
        subtitle_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(subtitle_label)

        main_layout.addSpacing(20)

        # Campos de Entrada
        self.user_input = QLineEdit()
        self.user_input.setPlaceholderText("Usuario")
        
        self.pass_input = QLineEdit()
        self.pass_input.setPlaceholderText("Contraseña")
        self.pass_input.setEchoMode(QLineEdit.Password)

        main_layout.addWidget(self.user_input)
        main_layout.addWidget(self.pass_input)

        # Label de error
        self.error_label = QLabel("")
        self.error_label.setObjectName("error_label")
        self.error_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.error_label)

        # Botón de Login
        btn_login = QPushButton("Login")
        btn_login.setObjectName("login_btn")
        btn_login.clicked.connect(self.intentar_login)
        main_layout.addWidget(btn_login)

        main_layout.addStretch()

        # Eventos de teclado
        self.pass_input.returnPressed.connect(self.intentar_login)
        self.user_input.returnPressed.connect(lambda: self.pass_input.setFocus())

    def intentar_login(self):
        user = self.user_input.text().strip()
        pwd = self.pass_input.text()
        
        if not user or not pwd:
            self.error_label.setText("Por favor, complete todos los campos.")
            return
            
        self.error_label.setText("")
        self.setCursor(Qt.WaitCursor)
        
        resultado = self.db.validar_credenciales(user, pwd)
        
        self.setCursor(Qt.ArrowCursor)
        
        if resultado:
            self.user_data = {
                "id": resultado[0],
                "username": resultado[1],
                "rol": resultado[2]
            }
            self.accept()
        else:
            self.error_label.setText("Usuario o contraseña incorrectos.")
            self.pass_input.clear()
            self.pass_input.setFocus()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # 1. Validación de Licencia Global
    acceso = obtener_estado_acceso_software()
    if not acceso["permitido"]:
        QMessageBox.critical(None, "Acceso Denegado", 
                             f"{acceso['mensaje']}\n\n{acceso['detalle']}")
        sys.exit(1)
    
    # 2. Conexión a DB y Login
    db = SupabaseConnector()
    
    # Verificar conectividad básica antes de mostrar login
    diag = db.diagnosticar()
    if not diag["ok"]:
        QMessageBox.critical(None, "Error de Sistema", 
                             f"No se puede conectar a la base de datos:\n{diag['msg']}")
        sys.exit(1)

    login_dialog = LoginDialog(db)
    if login_dialog.exec() == QDialog.Accepted:
        user_data = login_dialog.user_data
        window = GCDRYSApp(user_data)
        window.show()
        sys.exit(app.exec())
    else:
        sys.exit(0)
