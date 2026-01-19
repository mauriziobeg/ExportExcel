def classFactory(iface):
    from .export_excel import ExportExcel
    return ExportExcel(iface)
