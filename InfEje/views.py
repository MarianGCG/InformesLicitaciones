from django.shortcuts import render
from django.db.models import (
    Q,
    Case,
    When,
    Value,
    IntegerField,
    CharField,
)
from django.db.models.functions import Coalesce, Cast

from .models import (
    Lote,
    Empresa,
    RegistroLicitacion
)

from django.contrib import messages
from django.shortcuts import render, redirect
from .importadores import importar_excel
from django.http import HttpResponse, JsonResponse

from django.db.models import Q

from .exportar_pdf import (
    construir_pdf,
    obtener_registros_filtrados,
    limpiar_nombre_archivo,
)

def consultar(request):

    # ========================================================
    # DATOS PARA LOS FILTROS
    # ========================================================

    lotes = (
        Lote.objects
        .all()
        .order_by("-fecha_carga")
    )

    # Las empresas salen EXCLUSIVAMENTE de la tabla Empresa.
    # Como el CUIT es unique=True, no hay empresas duplicadas.
    empresas = (
        Empresa.objects
        .all()
        .order_by("nombre")
    )

    estados = (
        RegistroLicitacion.objects
        .exclude(estado__isnull=True)
        .exclude(estado="")
        .values_list("estado", flat=True)
        .distinct()
        .order_by("estado")
    )

    # ========================================================
    # CONSULTA BASE
    # ========================================================

    registros = RegistroLicitacion.objects.all()

    # ========================================================
    # FILTRO POR LOTE
    # ========================================================

    lotes_seleccionados = request.GET.getlist("lote")

    if lotes_seleccionados:
        registros = registros.filter(
            lote_id__in=lotes_seleccionados
        )

    # ========================================================
    # FILTRO POR EMPRESA
    #
    # El usuario selecciona una Empresa.
    # Obtenemos su CUIT.
    #
    # Luego buscamos SOLAMENTE por CUIT:
    #
    # cuit_oferente
    #       O
    # cuit_proveedor
    # ========================================================

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

    # ========================================================
    # FILTRO POR ESTADO
    #
    # Recordar:
    # Estado puede no existir en algunos Excel.
    # Por eso NO se utiliza para determinar si existe
    # una contratación.
    # ========================================================

    estado = request.GET.get("estado")

    if estado:

        registros = registros.filter(
            estado=estado
        )

    # ========================================================
    # FILTRO POR ORDEN DE COMPRA
    # ========================================================

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

    # ========================================================
    # FILTRO POR NÚMERO DE PROCESO
    # ========================================================

    proceso = request.GET.get("proceso")

    if proceso:

        registros = registros.filter(
            numero_proceso__icontains=proceso
        )

    # ========================================================
    # ORDEN DE LOS RESULTADOS
    # ========================================================
    registros = (
        registros
        .select_related(
            "lote",
            "empresa_oferente",
            "empresa_proveedor"
        )
        .annotate(
            empresa_orden=Coalesce(
                "empresa_oferente__nombre",
                "empresa_proveedor__nombre",
            ),

            tiene_oc=Case(
                When(
                    numero_oc__isnull=False,
                    then=Value(1)
                ),
                When(
                    numero_oc="",
                    then=Value(0)
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


    # ========================================================
    # CONTEXTO
    # ========================================================

    contexto = {

        "lotes": lotes,

        "empresas": empresas,

        "estados": estados,

        "registros": registros,

        "lotes_seleccionados": lotes_seleccionados,
        "empresa_seleccionada": empresa_id,

        "estado_seleccionado": estado,

        "oc_seleccionado": oc,

        "proceso_seleccionado": proceso,
    }

    return render(
        request,
        "InfEje/consultar.html",
        contexto
    )

def importar(request):

    if request.method == "POST":

        archivo = request.FILES.get("archivo")

        if not archivo:

            messages.error(
                request,
                "Debe seleccionar un archivo Excel."
            )

            return redirect("importar")

        if not archivo.name.lower().endswith(
            (".xlsx", ".xls")
        ):

            messages.error(
                request,
                "El archivo debe ser Excel (.xlsx o .xls)."
            )

            return redirect("importar")

        try:

            lote, cantidad, empresas_creadas = (
                importar_excel(
                    archivo,
                    archivo.name
                )
            )

            messages.success(
                request,
                (
                    f"Archivo importado correctamente. "
                    f"Lote: {lote.nombre_archivo}. "
                    f"Registros: {cantidad}. "
                    f"Empresas nuevas: {empresas_creadas}."
                )
            )

            return redirect("importar")

        except ValueError as error:

            messages.error(
                request,
                str(error)
            )

        except Exception as error:

            messages.error(
                request,
                f"Error durante la importación: {error}"
            )

    return render(
        request,
        "InfEje/importar.html"
    )

def empresas_para_pdf(request):

    registros, _ = obtener_registros_filtrados(request)

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

    return JsonResponse({
        "empresas": resultado
    })


def exportar_pdf_empresa(request, empresa_id):

    try:

        empresa = Empresa.objects.get(
            id=empresa_id
        )

    except Empresa.DoesNotExist:

        return HttpResponse(
            "Empresa no encontrada.",
            status=404,
        )

    registros, _ = obtener_registros_filtrados(request)

    registros = registros.filter(
        Q(cuit_oferente=empresa.cuit) |
        Q(cuit_proveedor=empresa.cuit)
    )

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

    respuesta["Content-Disposition"] = (
        f'inline; filename="{nombre_archivo}"'
    )

    return respuesta