from io import BytesIO

from django.http import HttpResponse

from django.db.models import (
    Q,
    Case,
    When,
    Value,
    IntegerField,
)

from django.db.models.functions import (
    Coalesce,
    Cast,
)
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_RIGHT
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
)

from .models import Empresa, RegistroLicitacion
from datetime import datetime
from reportlab.pdfbase.pdfmetrics import stringWidth
def agregar_encabezado_pagina(canvas, documento):
    canvas.saveState()

    fecha = datetime.now().strftime("%d/%m/%Y")

    texto_fecha = f"Fecha de emisión: {fecha}"
    texto_pagina = f"Página {documento.page}"

    ancho, alto = landscape(A4)

    canvas.setFont("Helvetica", 7)

    # Fecha de emisión
    canvas.drawRightString(
        ancho - 5 * mm - 28 * mm,
        alto - 5 * mm,
        texto_fecha
    )

    # Página
    canvas.drawRightString(
        ancho - 5 * mm,
        alto - 5 * mm,
        texto_pagina
    )

    canvas.restoreState()

def obtener_registros_filtrados(request):
    """
    Aplica exactamente los mismos filtros que utiliza
    la pantalla de consulta.
    """

    registros = RegistroLicitacion.objects.all()

    # --------------------------------------------------------
    # LOTES
    # --------------------------------------------------------

    lotes_seleccionados = request.GET.getlist("lote")

    if lotes_seleccionados:
        registros = registros.filter(
            lote_id__in=lotes_seleccionados
        )

    # --------------------------------------------------------
    # EMPRESA
    # --------------------------------------------------------

    empresa_id = request.GET.get("empresa")

    empresa_seleccionada = None

    if empresa_id:

        try:
            empresa_seleccionada = Empresa.objects.get(
                id=empresa_id
            )

            cuit = empresa_seleccionada.cuit

            registros = registros.filter(
                Q(cuit_oferente=cuit) |
                Q(cuit_proveedor=cuit)
            )

        except Empresa.DoesNotExist:
            pass

    # --------------------------------------------------------
    # ESTADO
    # --------------------------------------------------------

    estado = request.GET.get("estado")

    if estado:
        registros = registros.filter(
            estado=estado
        )

    # --------------------------------------------------------
    # ORDEN DE COMPRA
    # --------------------------------------------------------

    oc = request.GET.get("oc", "")

    if oc == "si":

        registros = registros.exclude(
            Q(numero_oc__isnull=True) |
            Q(numero_oc="")
        )

    elif oc == "no":

        registros = registros.filter(
            Q(numero_oc__isnull=True) |
            Q(numero_oc="")
        )

    # --------------------------------------------------------
    # PROCESO
    # --------------------------------------------------------

    proceso = request.GET.get("proceso")

    if proceso:

        registros = registros.filter(
            numero_proceso__icontains=proceso
        )

    # --------------------------------------------------------
    # ORDEN
    # --------------------------------------------------------

    registros = (
        registros
        .select_related(
            "lote",
            "empresa_oferente",
            "empresa_proveedor",
        )
        .annotate(
            empresa_orden=Coalesce(
                "empresa_oferente__nombre",
                "empresa_proveedor__nombre",
            ),

            tiene_oc=Case(
                When(
                    numero_oc__isnull=False,
                    then=Value(1),
                ),
                When(
                    numero_oc="",
                    then=Value(0),
                ),
                default=Value(0),
                output_field=IntegerField(),
            ),

            renglon_orden=Cast(
                "numero_renglon",
                IntegerField(),
            ),
        )
        .order_by(
            "empresa_orden",
            "-tiene_oc",
            "comprador",
            "numero_proceso",
            "renglon_orden",
        )
    )

    return registros, empresa_seleccionada


def formatear_importe(valor):
    if valor is None:
        return "-"

    try:
        valor = float(valor)
        texto = f"{valor:,.2f}"

        return (
            texto
            .replace(",", "@")
            .replace(".", ",")
            .replace("@", ".")
        )

    except (ValueError, TypeError):
        return str(valor)


def nombre_empresa(registro):

    if registro.empresa_oferente:
        return registro.empresa_oferente.nombre

    if registro.empresa_proveedor:
        return registro.empresa_proveedor.nombre

    if registro.oferente:
        return registro.oferente

    if registro.proveedor:
        return registro.proveedor

    return "-"


def situacion_empresa(registro):

    if registro.empresa_proveedor:
        return "Proveedor"

    if registro.empresa_oferente:
        return "Oferente"

    if registro.proveedor:
        return "Proveedor"

    if registro.oferente:
        return "Oferente"

    return "-"


def exportar_pdf(request):

    registros, empresa_seleccionada = (
        obtener_registros_filtrados(request)
    )

    buffer = BytesIO()

    documento = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=5  * mm,
        leftMargin=5  * mm,
        topMargin=8  * mm,
        bottomMargin=8  * mm,
        title="Consulta de Licitaciones",
    )

    estilos = getSampleStyleSheet()

    estilo_titulo = ParagraphStyle(
        "TituloConsulta",
        parent=estilos["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=16,
        leading=19,
        spaceAfter=5,
    )

    estilo_subtitulo = ParagraphStyle(
        "SubtituloConsulta",
        parent=estilos["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        textColor=colors.grey,
    )

    estilo_celda = ParagraphStyle(
        "Celda",
        parent=estilos["Normal"],
        fontName="Helvetica",
        fontSize=6.5,
        leading=8,
    )

    estilo_celda_negrita = ParagraphStyle(
        "CeldaNegrita",
        parent=estilos["Normal"],
        fontName="Helvetica-Bold",
        fontSize=6.5,
        leading=8,
    )

    estilo_importe = ParagraphStyle(
        "Importe",
        parent=estilo_celda,
        alignment=TA_RIGHT,
    )

    estilo_descripcion = ParagraphStyle(
        "Descripcion",
        parent=estilos["Normal"],
        fontName="Helvetica",
        fontSize=7.2,
        leading=8.7,
    )

    elementos = []

    elementos.append(
        Paragraph(
            "CONSULTA DE LICITACIONES",
            estilo_titulo,
        )
    )

    # --------------------------------------------------------
    # DESCRIPCIÓN DE FILTROS
    # --------------------------------------------------------

    filtros = []

    if empresa_seleccionada:

        filtros.append(
            f"Empresa: "
            f"{empresa_seleccionada.nombre} "
            f"- CUIT {empresa_seleccionada.cuit}"
        )

    lotes = request.GET.getlist("lote")

    if lotes:

        filtros.append(
            f"Lotes seleccionados: {len(lotes)}"
        )

    estado = request.GET.get("estado")

    if estado:
        filtros.append(f"Estado: {estado}")

    oc = request.GET.get("oc")

    if oc == "si":
        filtros.append("Orden de Compra: Con OC")

    elif oc == "no":
        filtros.append("Orden de Compra: Sin OC")

    proceso = request.GET.get("proceso")

    if proceso:
        filtros.append(
            f"Proceso: {proceso}"
        )

    if filtros:

        elementos.append(
            Paragraph(
                " | ".join(filtros),
                estilo_subtitulo,
            )
        )

    elementos.append(Spacer(1, 5 * mm))

    # --------------------------------------------------------
    # CABECERA TABLA
    # --------------------------------------------------------

    datos = [
        [
            Paragraph("<b>Empresa</b>", estilo_celda),
            Paragraph("<b>OC</b>", estilo_celda),
            Paragraph("<b>Comprador</b>", estilo_celda),
            Paragraph("<b>Proceso</b>", estilo_celda),
            Paragraph("<b>Renglón</b>", estilo_celda),
            Paragraph("<b>Estado</b>", estilo_celda),
            Paragraph("<b>Oferta</b>", estilo_celda),
            Paragraph("<b>Mon.</b>", estilo_celda),
            Paragraph("<b>Descripción</b>", estilo_celda),
        ]
    ]

    # --------------------------------------------------------
    # FILAS
    # --------------------------------------------------------

    for registro in registros:

        datos.append(
            [
                Paragraph(
                    nombre_empresa(registro),
                    estilo_celda_negrita,
                ),

                Paragraph(
                    registro.numero_oc or "-",
                    estilo_celda_negrita,
                ),

                Paragraph(
                    registro.comprador or "-",
                    estilo_celda,
                ),

                Paragraph(
                    registro.numero_proceso or "-",
                    estilo_celda_negrita,
                ),

                Paragraph(
                    str(
                        registro.numero_renglon
                        or "-"
                    ),
                    estilo_celda,
                ),


                Paragraph(
                    registro.estado or "-",
                    estilo_celda,
                ),

                Paragraph(
                    formatear_importe(
                        registro.precio_total_oferta
                    ),
                    estilo_importe,
                ),

                Paragraph(
                    registro.moneda_oferta or "-",
                    estilo_celda,
                ),


                Paragraph(
                    registro.descripcion_renglon or "-",
                    estilo_descripcion,
                ),
            ]
        )

    # --------------------------------------------------------
    # TABLA
    # --------------------------------------------------------

    tabla = Table(
        datos,
        colWidths=[


            25 * mm,   # Empresa
            23 * mm,   # OC
            38 * mm,   # Comprador
            27 * mm,   # Proceso
            10 * mm,   # Renglón
            20 * mm,   # Estado
            22 * mm,   # Oferta
            12 * mm,   # Moneda
            105 * mm,  # Descripción

        ],
        repeatRows=1,
    )

    tabla.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor("#E9ECEF"),
                ),
                (
                    "TEXTCOLOR",
                    (0, 0),
                    (-1, 0),
                    colors.black,
                ),
                (
                    "FONTNAME",
                    (0, 0),
                    (-1, 0),
                    "Helvetica-Bold",
                ),
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE",
                ),
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.3,
                    colors.HexColor("#D9D9D9"),
                ),
                (
                    "ROWBACKGROUNDS",
                    (0, 1),
                    (-1, -1),
                    [
                        colors.white,
                        colors.HexColor("#F8F9FA"),
                    ],
                ),
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    4,
                ),
                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    4,
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    3,
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    3,
                ),
            ]
        )
    )

    elementos.append(tabla)

    elementos.append(
        Spacer(1, 4 * mm)
    )

    elementos.append(
        Paragraph(
            f"Total de registros: {registros.count()}",
            estilo_subtitulo,
        )
    )

    documento.build(
        elementos,
        onFirstPage=agregar_encabezado_pagina,
        onLaterPages=agregar_encabezado_pagina,
    )
    buffer.seek(0)

    respuesta = HttpResponse(
        buffer,
        content_type="application/pdf",
    )

    respuesta["Content-Disposition"] = (
        'inline; filename="consulta_licitaciones.pdf"'
    )

    return respuesta