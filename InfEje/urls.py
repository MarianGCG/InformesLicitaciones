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

]