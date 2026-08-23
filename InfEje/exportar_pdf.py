from io import BytesIO
from datetime import datetime
import re

from django.http import HttpResponse, JsonResponse
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
from reportlab.lib.styles import (
    getSampleStyleSheet,
    ParagraphStyle,
)
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
)

from .models import (
    Empresa,
    RegistroLicitacion,
)


# ============================================================
# ENCABEZADO DE PÁGINA
# ============================================================

def agregar_encabezado_pagina(canvas, documento):

    canvas.saveState()

    fecha = datetime.now().strftime("%d/%m/%Y")

    texto_fecha = f"Fecha de emisión: {fecha}"
    texto_pagina = f"Página {documento.page}"

    ancho, alto = landscape(A4)

    canvas.setFont("Helvetica", 7)

    canvas.drawRightString(
        ancho - 5 * mm - 28 * mm,
        alto - 5 * mm,
        texto_fecha,
    )

    canvas.drawRightString(
        ancho - 5 * mm,
        alto - 5 * mm,
        texto_pagina,
    )

    canvas.restoreState()


# ============================================================
# FILTROS
# ============================================================

def obtener_registros_filtrados(request):

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

            empresa_seleccionada = None

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


# ============================================================
# UTILIDADES
# ============================================================

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


def limpiar_nombre_archivo(nombre):

    # Sacar puntos
    nombre = nombre.replace(".", "")

    # Caracteres que Windows no permite
    nombre = re.sub(
        r'[\\/:*?"<>|]',
        "",
        nombre,
    )

    # Espacios repetidos
    nombre = re.sub(
        r"\s+",
        " ",
        nombre,
    ).strip()

    return nombre or "Empresa"


# ============================================================
# CREAR PDF
# ============================================================

def construir_pdf(
    registros,
    empresa=None,
):

    buffer = BytesIO()

    documento = SimpleDocTemplate(
        buffer,

        pagesize=landscape(A4),

        rightMargin=8 * mm,
        leftMargin=8 * mm,

        topMargin=8 * mm,
        bottomMargin=8 * mm,

        title="Consulta de Licitaciones",
    )

    estilos = getSampleStyleSheet()

    # --------------------------------------------------------
    # ESTILOS
    # --------------------------------------------------------

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
        fontSize=6.8,
        leading=8,
    )

    estilo_celda_negrita = ParagraphStyle(
        "CeldaNegrita",
        parent=estilos["Normal"],
        fontName="Helvetica-Bold",
        fontSize=6.8,
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
        fontSize=7.0,
        leading=8.3,
    )

    estilo_encabezado = ParagraphStyle(
        "Encabezado",
        parent=estilos["Normal"],
        fontName="Helvetica-Bold",
        fontSize=6.5,
        leading=7,
    )

    elementos = []

    # --------------------------------------------------------
    # TÍTULO
    # --------------------------------------------------------

    elementos.append(
        Paragraph(
            "CONSULTA DE LICITACIONES",
            estilo_titulo,
        )
    )

    # --------------------------------------------------------
    # EMPRESA
    # --------------------------------------------------------

    if empresa:

        elementos.append(
            Paragraph(
                (
                    f"Empresa: {empresa.nombre} "
                    f"- CUIT {empresa.cuit}"
                ),
                estilo_subtitulo,
            )
        )

    # --------------------------------------------------------
    # TABLA
    # --------------------------------------------------------

    datos = [
        [
            Paragraph(
                "<b>Empresa</b>",
                estilo_encabezado,
            ),

            Paragraph(
                "<b>OC</b>",
                estilo_encabezado,
            ),

            Paragraph(
                "<b>Comprador</b>",
                estilo_encabezado,
            ),

            Paragraph(
                "<b>Proceso</b>",
                estilo_encabezado,
            ),

            Paragraph(
                "<b>Renglón</b>",
                estilo_encabezado,
            ),

            Paragraph(
                "<b>Apertura</b>",
                estilo_encabezado,
            ),

            Paragraph(
                "<b>Inicio</b>",
                estilo_encabezado,
            ),

            Paragraph(
                "<b>Fin</b>",
                estilo_encabezado,
            ),

            Paragraph(
                "<b>Estado</b>",
                estilo_encabezado,
            ),

            Paragraph(
                "<b>Oferta</b>",
                estilo_encabezado,
            ),

            Paragraph(
                "<b>Mon.</b>",
                estilo_encabezado,
            ),

            Paragraph(
                "<b>Descripción</b>",
                estilo_encabezado,
            ),
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
                    (
                        registro.fecha_apertura.strftime(
                            "%d/%m/%Y"
                        )
                        if registro.fecha_apertura
                        else "-"
                    ),
                    estilo_celda,
                ),

                Paragraph(
                    (
                        registro.fecha_inicio_contrato.strftime(
                            "%d/%m/%Y"
                        )
                        if registro.fecha_inicio_contrato
                        else "-"
                    ),
                    estilo_celda,
                ),

                Paragraph(
                    (
                        registro.fecha_fin_contrato.strftime(
                            "%d/%m/%Y"
                        )
                        if registro.fecha_fin_contrato
                        else "-"
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
            30 * mm,   # Empresa
            22 * mm,   # OC
            38 * mm,   # Comprador
            24 * mm,   # Proceso
            9 * mm,    # Renglón
            15 * mm,   # Apertura
            15 * mm,   # Inicio
            15 * mm,   # Fin
            17 * mm,   # Estado
            20 * mm,   # Oferta
            10 * mm,   # Moneda
            65 * mm,   # Descripción
        ],

        repeatRows=1,

        hAlign="CENTER",
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
                    2.5,
                ),

                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    2.5,
                ),

                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    2.5,
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    2.5,
                ),
            ]
        )
    )

    elementos.append(
        Spacer(1, 5 * mm)
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

    # --------------------------------------------------------
    # CREAR DOCUMENTO
    # --------------------------------------------------------

    documento.build(
        elementos,

        onFirstPage=agregar_encabezado_pagina,

        onLaterPages=agregar_encabezado_pagina,
    )

    buffer.seek(0)

    return buffer


# ============================================================
# PDF GENERAL
# ============================================================

def exportar_pdf(request):

    registros, empresa_seleccionada = (
        obtener_registros_filtrados(request)
    )

    buffer = construir_pdf(
        registros,
        empresa=empresa_seleccionada,
    )

    respuesta = HttpResponse(
        buffer.getvalue(),
        content_type="application/pdf",
    )

    respuesta[
        "Content-Disposition"
    ] = (
        'inline; '
        'filename="consulta_licitaciones.pdf"'
    )

    return respuesta


# ============================================================
# LISTA DE EMPRESAS DE LA CONSULTA
# ============================================================

def empresas_para_pdf(request):

    registros, _ = obtener_registros_filtrados(
        request
    )

    empresas = {}

    for registro in registros:

        empresa = None

        if registro.empresa_oferente:

            empresa = registro.empresa_oferente

        elif registro.empresa_proveedor:

            empresa = registro.empresa_proveedor

        if not empresa:
            continue

        empresas[empresa.id] = {
            "id": empresa.id,
            "nombre": empresa.nombre,
            "cuit": empresa.cuit,
        }

    resultado = sorted(
        empresas.values(),
        key=lambda x: x["nombre"].lower(),
    )

    return JsonResponse(
        {
            "empresas": resultado
        }
    )


# ============================================================
# PDF INDIVIDUAL POR EMPRESA
# ============================================================

def exportar_pdf_empresa(
    request,
    empresa_id,
):

    try:

        empresa = Empresa.objects.get(
            id=empresa_id
        )

    except Empresa.DoesNotExist:

        return HttpResponse(
            "Empresa no encontrada.",
            status=404,
        )

    # --------------------------------------------------------
    # APLICAR LOS MISMOS FILTROS DE LA CONSULTA
    # --------------------------------------------------------

    registros, _ = obtener_registros_filtrados(
        request
    )

    # --------------------------------------------------------
    # FILTRAR POR CUIT
    # --------------------------------------------------------

    registros = registros.filter(
        Q(cuit_oferente=empresa.cuit) |
        Q(cuit_proveedor=empresa.cuit)
    )

    # --------------------------------------------------------
    # GENERAR PDF
    # --------------------------------------------------------

    buffer = construir_pdf(
        registros,
        empresa=empresa,
    )

    nombre_archivo = (
        limpiar_nombre_archivo(
            empresa.nombre
        )
        + ".pdf"
    )

    respuesta = HttpResponse(
        buffer.getvalue(),
        content_type="application/pdf",
    )

    respuesta[
        "Content-Disposition"
    ] = (
        'inline; '
        f'filename="{nombre_archivo}"'
    )

    return respuesta