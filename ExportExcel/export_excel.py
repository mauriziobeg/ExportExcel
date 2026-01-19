import os
import shutil
import tempfile
import datetime

from qgis.PyQt.QtGui import QIcon
from qgis.PyQt.QtWidgets import (
    QAction, QDialog, QVBoxLayout, QHBoxLayout,
    QListWidget, QListWidgetItem, QPushButton,
    QFileDialog, QMessageBox, QComboBox, QInputDialog
)
from qgis.PyQt.QtCore import Qt
from qgis.utils import iface
from qgis.core import QgsVectorFileWriter, QgsProject, QgsSettings


class ExportExcel:

    def __init__(self, iface):
        self.iface = iface
        self.action = None

    # --------------------------------------------------
    # GUI
    # --------------------------------------------------
    def initGui(self):
        plugin_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(plugin_dir, "icon.png")

        if os.path.exists(icon_path):
            self.action = QAction(QIcon(icon_path), "Export Excel", self.iface.mainWindow())
        else:
            self.action = QAction("Export Excel", self.iface.mainWindow())

        self.action.triggered.connect(self.run)
        self.iface.addToolBarIcon(self.action)
        self.iface.addPluginToMenu("Export Excel", self.action)

    def unload(self):
        self.iface.removeToolBarIcon(self.action)
        self.iface.removePluginMenu("Export Excel", self.action)

    # --------------------------------------------------
    # MAIN
    # --------------------------------------------------
    def run(self):

        layer = self.iface.activeLayer()
        if not layer or layer.selectedFeatureCount() == 0:
            self.iface.messageBar().pushCritical(
                "Errore", "Layer non valido o nessuna feature selezionata"
            )
            return

        settings = QgsSettings()
        base_key = f"export_excel/presets/{layer.id()}"
        preset_names = settings.value(f"{base_key}/names", [], type=list)

        dialog = QDialog(self.iface.mainWindow())
        dialog.setWindowTitle("Export Excel")
        layout = QVBoxLayout(dialog)

        # ---- Preset
        preset_combo = QComboBox()
        preset_combo.addItem("<Nuovo preset>")
        preset_combo.addItems(preset_names)
        layout.addWidget(preset_combo)

        # ---- Lista campi
        list_widget = QListWidget()
        list_widget.setDragDropMode(QListWidget.InternalMove)
        layout.addWidget(list_widget)

        fields = {f.name(): f for f in layer.fields()}

        def load_fields(field_names):
            list_widget.clear()
            for fname in field_names:
                f = fields[fname]
                alias = f.alias() or f.name()
                label = f"{alias} ({f.name()})"
                item = QListWidgetItem(label)
                item.setData(Qt.UserRole, f.name())
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Checked)
                list_widget.addItem(item)

        def preset_changed():
            idx = preset_combo.currentIndex()
            if idx == 0:
                load_fields(list(fields.keys()))
            else:
                pname = preset_combo.currentText()
                saved = settings.value(f"{base_key}/{pname}", [], type=list)
                load_fields(saved)

        preset_combo.currentIndexChanged.connect(preset_changed)
        preset_changed()

        # ---- Pulsanti
        btns = QHBoxLayout()
        btn_export = QPushButton("Esporta")
        btn_save = QPushButton("Salva preset")
        btns.addWidget(btn_export)
        btns.addWidget(btn_save)
        layout.addLayout(btns)

        # --------------------------------------------------
        # SALVA PRESET
        # --------------------------------------------------
        def save_preset():
            name, ok = QInputDialog.getText(dialog, "Preset", "Nome preset:")
            if not ok or not name:
                return

            fields_sel = [
                list_widget.item(i).data(Qt.UserRole)
                for i in range(list_widget.count())
                if list_widget.item(i).checkState() == Qt.Checked
            ]

            if not fields_sel:
                QMessageBox.warning(dialog, "Errore", "Nessun campo selezionato")
                return

            if name not in preset_names:
                preset_names.append(name)
                settings.setValue(f"{base_key}/names", preset_names)
                preset_combo.addItem(name)

            settings.setValue(f"{base_key}/{name}", fields_sel)
            preset_combo.setCurrentText(name)

        btn_save.clicked.connect(save_preset)

        # --------------------------------------------------
        # EXPORT (FILE TEMPORANEO → COPIA)
        # --------------------------------------------------
        def do_export():

            selected_fields = [
                list_widget.item(i).data(Qt.UserRole)
                for i in range(list_widget.count())
                if list_widget.item(i).checkState() == Qt.Checked
            ]

            if not selected_fields:
                QMessageBox.warning(dialog, "Errore", "Nessun campo selezionato")
                return

            last_dir = settings.value("export_excel/last_dir", os.path.expanduser("~"))
            ts = datetime.datetime.now().strftime("%Y%m%d")

            final_path, _ = QFileDialog.getSaveFileName(
                dialog,
                "Salva Excel",
                os.path.join(last_dir, f"exportQGis{ts}.xlsx"),
                "Excel (*.xlsx)"
            )

            if not final_path:
                return

            settings.setValue("export_excel/last_dir", os.path.dirname(final_path))

            # ---- FILE TEMPORANEO
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            tmp_path = tmp.name
            tmp.close()

            options = QgsVectorFileWriter.SaveVectorOptions()
            options.driverName = "XLSX"
            options.onlySelectedFeatures = True
            options.includeGeometry = False
            options.attributes = [
                layer.fields().indexFromName(f) for f in selected_fields
            ]

            res = QgsVectorFileWriter.writeAsVectorFormatV2(
                layer,
                tmp_path,
                QgsProject.instance().transformContext(),
                options
            )

            if isinstance(res, tuple):
                error_code, error_msg = res[0], res[1]
            else:
                error_code, error_msg = res, ""

            if error_code != QgsVectorFileWriter.NoError:
                os.remove(tmp_path)
                self.iface.messageBar().pushCritical(
                    "Errore", f"Export XLSX fallito: {error_msg}"
                )
                return

            # ---- COPIA → FILE FINALE
            shutil.copy2(tmp_path, final_path)
            os.remove(tmp_path)

            self.iface.messageBar().pushSuccess(
                "OK", f"Excel creato correttamente ({len(selected_fields)} campi)"
            )
            dialog.accept()

        btn_export.clicked.connect(do_export)
        dialog.exec()
