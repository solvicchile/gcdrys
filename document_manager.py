from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem, 
    QHeaderView, QComboBox, QLineEdit, QPushButton, QLabel, QFrame,
    QMessageBox
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

class DocumentManager(QWidget):
    def __init__(self, db_connector, user_data, parent=None):
        super().__init__(parent)
        self.db = db_connector
        self.user_data = user_data
        self.current_documents = []
        
        self.init_ui()
        self.load_filters()
        self.refresh_table()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # --- Panel de Filtros ---
        filter_frame = QFrame()
        filter_frame.setStyleSheet("background: #f9f9f9; border-bottom: 1px solid #dcdcdc;")
        filter_layout = QHBoxLayout(filter_frame)
        
        # Filtro Área
        filter_layout.addWidget(QLabel("Área:"))
        self.combo_area = QComboBox()
        self.combo_area.addItem("TODAS")
        filter_layout.addWidget(self.combo_area)
        
        # Filtro Tipo
        filter_layout.addWidget(QLabel("Tipo:"))
        self.combo_tipo = QComboBox()
        self.combo_tipo.addItem("TODOS")
        filter_layout.addWidget(self.combo_tipo)
        
        # Filtro Estado
        filter_layout.addWidget(QLabel("Estado:"))
        self.combo_estado = QComboBox()
        self.combo_estado.addItem("TODOS")
        # Estados estándar del sistema
        estados = ["Borrador", "En revision", "Pendiente aprobacion", "Vigente", "Obsoleto", "Anulado"]
        self.combo_estado.addItems(estados)
        filter_layout.addWidget(self.combo_estado)
        
        # Búsqueda por texto
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Buscar por nombre o código...")
        filter_layout.addWidget(self.search_input, stretch=1)
        
        # Botón Buscar
        btn_search = QPushButton("Filtrar")
        btn_search.clicked.connect(self.refresh_table)
        filter_layout.addWidget(btn_search)
        
        layout.addWidget(filter_frame)

        # --- Tabla de Documentos ---
        self.table = QTableWidget()
        self.table.setColumnCount(8)
        headers = ["ID", "Código", "Nombre", "Área", "Tipo", "Versión", "Estado", "Confidencial"]
        self.table.setHorizontalHeaderLabels(headers)
        
        # Ajuste de columnas
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(2, QHeaderView.Stretch) # Nombre ocupa el espacio restante
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents) # ID ajustado
        
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setAlternatingRowColors(True)
        
        layout.addWidget(self.table)

        # --- Barra de Acciones Inferior ---
        action_frame = QFrame()
        action_layout = QHBoxLayout(action_frame)
        action_layout.setContentsMargins(10, 5, 10, 5)
        
        self.btn_edit = QPushButton("Editar Documento")
        self.btn_edit.clicked.connect(self.on_edit_clicked)
        # El estado del botón se actualizará según la selección y permisos
        self.btn_edit.setEnabled(False) 
        
        action_layout.addWidget(self.btn_edit)
        action_layout.addStretch()
        
        layout.addWidget(action_frame)
        
        # Conexión para habilitar/deshabilitar botón al seleccionar fila
        self.table.itemSelectionChanged.connect(self.on_selection_changed)

    def load_filters(self):
        """Carga las listas de Áreas y Tipos desde la DB."""
        areas = self.db.obtener_opciones_filtro("areas", "nombre")
        for area in areas:
            self.combo_area.addItem(area)
            
        tipos = self.db.obtener_opciones_filtro("tipos_documento", "nombre")
        for tipo in tipos:
            self.combo_tipo.addItem(tipo)

    def refresh_table(self):
        """Obtiene datos de la DB y llena la tabla."""
        area_val = None if self.combo_area.currentText() == "TODAS" else self.combo_area.currentText()
        tipo_val = None if self.combo_tipo.currentText() == "TODOS" else self.combo_tipo.currentText()
        estado_val = None if self.combo_estado.currentText() == "TODOS" else self.combo_estado.currentText()
        texto_val = self.search_input.text().strip() if self.search_input.text().strip() else None
        
        docs = self.db.obtener_documentos(area=area_val, tipo=tipo_val, estado=estado_val, texto=texto_val)
        self.current_documents = docs
        
        self.table.setRowCount(0)
        for doc in docs:
            row_pos = self.table.rowCount()
            self.table.insertRow(row_pos)
            
            # Mapeo de columnas: id, codigo, nombre, area, tipo, version, estado, confidencial
            self.table.setItem(row_pos, 0, QTableWidgetItem(str(doc['id'])))
            self.table.setItem(row_pos, 1, QTableWidgetItem(doc['codigo']))
            self.table.setItem(row_pos, 2, QTableWidgetItem(doc['nombre']))
            self.table.setItem(row_pos, 3, QTableWidgetItem(doc['area']))
            self.table.setItem(row_pos, 4, QTableWidgetItem(doc['tipo']))
            self.table.setItem(row_pos, 5, QTableWidgetItem(str(doc['version'])))
            self.table.setItem(row_pos, 6, QTableWidgetItem(doc['estado']))
            
            conf_item = QTableWidgetItem("Sí" if doc['confidencial'] else "No")
            if doc['confidencial']:
                conf_item.setForeground(Qt.red) # Resaltar confidencial
            self.table.setItem(row_pos, 7, conf_item)

    def on_selection_changed(self):
        """Habilita el botón de editar si hay selección y permisos."""
        selected = self.table.selectedItems()
        if not selected:
            self.btn_edit.setEnabled(False)
            return
            
        # Verificar permisos básicos (ejemplo simplificado)
        # En una implementación completa, verificaríamos el rol y el estado del doc seleccionado
        if self.user_data['rol'] in ['admin', 'Superusuario', 'Gestor']:
            self.btn_edit.setEnabled(True)
        else:
            self.btn_edit.setEnabled(False)

    def on_edit_clicked(self):
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            return
            
        row = selected_rows[0].row()
        doc_id = self.table.item(row, 0).text()
        doc_estado = self.table.item(row, 6).text()
        
        # Aquí iría la lógica para abrir el diálogo de edición
        # Por ahora, solo mostramos un mensaje
        QMessageBox.information(self, "Editar", f"Iniciando edición para documento ID: {doc_id}\nEstado actual: {doc_estado}")
