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
]