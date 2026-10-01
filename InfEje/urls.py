from django.urls import path

from . import views
from .exportar_pdf import exportar_pdf


urlpatterns = [

    path(
        "",
        views.consultar,
        name="consultar"
    ),

    path(
        "importar/",
        views.importar,
        name="importar"
    ),

    path(
        "exportar-pdf/",
        exportar_pdf,
        name="exportar_pdf"
    ),

    path(
        "empresas-para-pdf/",
        views.empresas_para_pdf,
        name="empresas_para_pdf",
    ),

    path(
        "exportar-pdf-empresa/<int:empresa_id>/",
        views.exportar_pdf_empresa,
        name="exportar_pdf_empresa",
    ),


    path(
        "empresas/",
        views.empresas,
        name="empresas",
    ),

    path(
        "importar-emails/",
        views.importar_emails,
        name="importar_emails",
    ),

    path(
        "empresas/<int:empresa_id>/guardar-emails/",
        views.guardar_emails,
        name="guardar_emails",
    ),


    path(
        "exportar-csv-doppler/",
        views.exportar_csv_doppler,
        name="exportar_csv_doppler",
    ),
    path(
        "guardar-ficha-empresa/<int:empresa_id>/",
        views.guardar_ficha_empresa,
        name="guardar_ficha_empresa",
    ),
    path(
        "lotes/<int:lote_id>/eliminar/",
        views.eliminar_lote,
        name="eliminar_lote",
    ),
    path(
        "exportar-empresas/",
        views.exportar_empresas,
        name="exportar_empresas",
    ),
        
    path(
        "exportar-word-mails/",
        views.exportar_word_mails,
        name="exportar_word_mails",
    ),

    path(
        "seguimiento-cartas/",
        views.seguimiento_cartas,
        name="seguimiento_cartas",
    ),

    path(
        "exportar-excel-seguimiento-cartas/",
        views.exportar_excel_seguimiento_cartas,
        name="exportar_excel_seguimiento_cartas"
    ),

    path(
        "editar-carta/",
        views.editar_carta,
        name="editar_carta",
    ),

    path(
        "api/plantilla-carta/<str:carta>/<str:tipo_destinatario>/",
        views.obtener_plantilla_carta,
        name="obtener_plantilla_carta",
    ),


    path(
        "api/guardar-plantilla-carta/",
        views.guardar_plantilla_carta,
        name="guardar_plantilla_carta",
    ),


    path(
        "api/aplicar-cartas-enviadas/",
        views.aplicar_cartas_enviadas,
    ),

    path(
        "api/guardar-seguimiento-carta/",
        views.guardar_seguimiento_carta,
        name="guardar_seguimiento_carta",
    ),


    path(
        "api/eliminar-seguimiento-carta/",
        views.eliminar_seguimiento_carta,
        name="eliminar_seguimiento_carta",
    ),


]