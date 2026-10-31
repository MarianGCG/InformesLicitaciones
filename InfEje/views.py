from django.shortcuts import render, redirect
from .models import SeguimientoCartasFlow
from django.db.models import (
    Q,
    Case,
    When,
    Value,
    IntegerField,
    CharField,
)
from django.db import transaction
from django.utils.dateparse import parse_date
import json
from django.db.models.functions import Coalesce, Cast, Lower

from .models import (
    Lote,
    Empresa,
    SeguimientoCartasFlow,
    CartaPlantilla,
    CartaDefinicion,
    CartaCondicion,
    RegistroLicitacion,
)
import csv
import re
from openpyxl import Workbook
from decimal import Decimal
from datetime import datetime, date
from django.contrib import messages
from django.shortcuts import render, redirect, get_object_or_404
from django.urls import reverse
from .importadores import importar_excel
from django.http import HttpResponse, JsonResponse

from django.views.decorators.http import require_POST, require_GET
from django.views.decorators.clickjacking import xframe_options_sameorigin
from collections import defaultdict
from django.db.models import Q, Count
from .exportar_pdf import (
    construir_pdf,
    obtener_registros_filtrados,
    limpiar_nombre_archivo,
)
from .importadores import (
    importar_excel,
    importar_emails_excel,
)
from .exportar_word import generar_word_mails



# ============================================================
# GRÁFICO DE VENCIMIENTOS DE CONTRATOS
# ============================================================

def grafico_vencimientos(request):

    # --------------------------------------------------------
    # 1. OBTENER REGISTROS ADJUDICADOS CON FECHA DE VENCIMIENTO
    # --------------------------------------------------------

    registros = (
        RegistroLicitacion.objects
        .filter(
            fecha_fin_contrato__isnull=False
        )
        .exclude(
            numero_oc__isnull=True
        )
        .exclude(
            numero_oc=""
        )
        .exclude(numero_proceso__isnull=True)
        .exclude(numero_proceso="")
        .exclude(numero_renglon__isnull=True)
        .exclude(numero_renglon="")
        .exclude(fecha_apertura__isnull=True)
        .select_related(
            "empresa_oferente",
            "empresa_proveedor",
        )
        .order_by("fecha_fin_contrato", "id")
    )

    # --------------------------------------------------------
    # 2. UNA SOLA FILA ADJUDICADA POR PROCESO + RENGLÓN
    #    + FECHA DE APERTURA
    # --------------------------------------------------------

    registros_unicos = {}

    for registro in registros:

        clave = (
            registro.numero_proceso,
            registro.numero_renglon,
            registro.fecha_apertura,
        )

        if clave not in registros_unicos:
            registros_unicos[clave] = registro

    # --------------------------------------------------------
    # 3. BUSCAR TODOS LOS REGISTROS DE ESOS PROCESOS
    #    INCLUYENDO LOS QUE NO TIENEN OC NI FECHA DE FIN
    # --------------------------------------------------------

    procesos = {
        clave[0]
        for clave in registros_unicos.keys()
    }

    todos_los_registros = (
        RegistroLicitacion.objects
        .filter(
            numero_proceso__in=procesos
        )
        .exclude(numero_renglon__isnull=True)
        .exclude(numero_renglon="")
        .exclude(fecha_apertura__isnull=True)
        .select_related(
            "empresa_oferente",
            "empresa_proveedor",
        )
    )

    registros_por_clave = {}

    for registro in todos_los_registros:

        clave = (
            registro.numero_proceso,
            registro.numero_renglon,
            registro.fecha_apertura,
        )

        if clave not in registros_unicos:
            continue

        registros_por_clave.setdefault(
            clave, []
        ).append(registro)

    # --------------------------------------------------------
    # 4. AGRUPAR POR MES DE VENCIMIENTO
    # --------------------------------------------------------

    vencimientos = defaultdict(list)

    empresa_seleccionada = request.GET.get("empresa", "")

    for clave, registro_adjudicado in registros_unicos.items():

        participantes = registros_por_clave.get(
            clave, []
        )

        # ----------------------------------------------------
        # REUNIR TODAS LAS EMPRESAS PARTICIPANTES
        # ----------------------------------------------------

        empresas = {}

        for participante in participantes:

            for empresa in (
                participante.empresa_oferente,
                participante.empresa_proveedor,
            ):

                if empresa:
                    empresas[empresa.id] = empresa

        # ----------------------------------------------------
        # FILTRO POR EMPRESA:
        # INCLUYE ADJUDICADOS Y NO ADJUDICADOS
        # ----------------------------------------------------

        if empresa_seleccionada:

            if str(empresa_seleccionada) not in [
                str(empresa_id)
                for empresa_id in empresas.keys()
            ]:
                continue

 

        # ----------------------------------------------------
        # FECHA DE VENCIMIENTO: SIEMPRE LA DEL ADJUDICADO
        # ----------------------------------------------------

        fecha = registro_adjudicado.fecha_fin_contrato

        periodo = fecha.strftime("%Y-%m")

        # Empresa adjudicada
        empresa_adjudicada = (
            registro_adjudicado.empresa_proveedor
            or registro_adjudicado.empresa_oferente
        )

        # Una fila por cada empresa participante
        for empresa in sorted(
            empresas.values(),
            key=lambda e: e.nombre
        ):

            adjudicado = "-"

            if empresa_adjudicada:

                if empresa.id == empresa_adjudicada.id:
                    adjudicado = "SI"
                else:
                    adjudicado = empresa_adjudicada.nombre

            vencimientos[periodo].append({

                "comprador": registro_adjudicado.comprador or "",

                "proceso": registro_adjudicado.numero_proceso,

                "renglon": registro_adjudicado.numero_renglon,

                "fecha_fin": fecha.strftime("%d/%m/%Y"),

                "proveedor": empresa.nombre,

                "adjudicado": adjudicado,

            })




    # --------------------------------------------------------
    # 5. PREPARAR DATOS PARA EL GRÁFICO
    # --------------------------------------------------------

    periodos = sorted(vencimientos.keys())

    etiquetas = [
        datetime.strptime(
            periodo,
            "%Y-%m"
        ).strftime("%m/%Y")
        for periodo in periodos
    ]

    cantidades = [
        len(vencimientos[periodo])
        for periodo in periodos
    ]

    return render(
        request,
        "InfEje/grafico_vencimientos.html",
        {
            "etiquetas": json.dumps(etiquetas),

            "cantidades": json.dumps(cantidades),

            "detalles": json.dumps([
                vencimientos[periodo]
                for periodo in periodos
            ], ensure_ascii=False),

        }
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
        .annotate(
            nombre_orden=Lower("nombre")
        )
        .order_by("nombre_orden")
    )

    archivos_iniciales = (
        Empresa.objects
        .exclude(archivo_inicial__isnull=True)
        .exclude(archivo_inicial="")
        .values_list("archivo_inicial", flat=True)
        .distinct()
        .order_by("-archivo_inicial")
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

    archivos_iniciales_seleccionados = request.GET.getlist(
        "archivo_inicial"
    )

    todas_archivos = (
        request.GET.get("todas_archivos", "1") == "1"
    )

    # ========================================================
    # FILTRO POR LOTE
    # ========================================================

    lotes_seleccionados = request.GET.getlist("lote")

    if lotes_seleccionados:
        registros = registros.filter(
            lote_id__in=lotes_seleccionados
        )

    # ========================================================
    # FILTRO POR archivos_iniciales_seleccionados
    # ========================================================

    if not todas_archivos:
        if archivos_iniciales_seleccionados:
            registros = registros.filter(
                Q(
                    empresa_oferente__archivo_inicial__in=
                    archivos_iniciales_seleccionados
                )
                |
                Q(
                    empresa_proveedor__archivo_inicial__in=
                    archivos_iniciales_seleccionados
                )
            )
        else:
            registros = registros.none()


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
                id=int(empresa_id)
            )

            cuit = empresa_seleccionada.cuit

            registros = registros.filter(
                Q(cuit_oferente=cuit)
                |
                Q(cuit_proveedor=cuit)
            )

        except (
            ValueError,
            TypeError,
            Empresa.DoesNotExist,
        ):

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

    elif oc:

        # Cuando el filtro trae una OC concreta (por ejemplo,
        # seleccionada mediante doble clic), buscar esa OC exacta.
        registros = registros.filter(
            numero_oc=oc
        )

    # ========================================================
    # FILTRO POR NÚMERO DE PROCESO
    # ========================================================

    proceso = request.GET.get("proceso")

    if proceso:

        registros = registros.filter(
            numero_proceso__icontains=proceso
        )

    # IMPORTANTE: la segunda pestaña usa todos los filtros de entrada
    # salvo el filtro Con OC / Sin OC. Ese filtro es exclusivo de la
    # tabla Resultados, como consulta rápida.
    registros_para_totales = registros

    # ========================================================
    # EXCLUIR EN RESULTADOS PROCESOS + RENGLONES
    # ADJUDICADOS A OTRA EMPRESA
    #
    # Solo afecta la pestaña RESULTADOS.
    # ========================================================

    mostrar_adjudicados = (
        request.GET.get("mostrar_adjudicados", "0") == "1"
    )

    ids_adjudicados_otra_empresa = []

    # Detectamos qué procesos/renglones tienen alguna adjudicación.
    adjudicaciones = (
        RegistroLicitacion.objects
        .exclude(
            Q(numero_oc__isnull=True) |
            Q(numero_oc="")
        )
        .exclude(numero_proceso__isnull=True)
        .exclude(numero_proceso="")
        .exclude(numero_renglon__isnull=True)
        .exclude(numero_renglon="")
        .values_list(
            "numero_proceso",
            "numero_renglon",
            "cuit_oferente",
            "cuit_proveedor",
        )
    )

    empresas_con_oc = {}

    for (
        numero_proceso,
        numero_renglon,
        cuit_oferente,
        cuit_proveedor,
    ) in adjudicaciones:

        clave = (
            numero_proceso,
            numero_renglon,
        )

        cuit_empresa = (
            cuit_oferente
            or cuit_proveedor
        )

        if not cuit_empresa:
            continue

        empresas_con_oc.setdefault(
            clave,
            set()
        ).add(cuit_empresa)


    # Revisamos solamente los registros que ya quedaron
    # dentro de la búsqueda actual.
    ids_a_excluir = []

    registros_actuales = registros.values_list(
        "id",
        "numero_proceso",
        "numero_renglon",
        "numero_oc",
        "cuit_oferente",
        "cuit_proveedor",
    )

    for (
        registro_id,
        numero_proceso,
        numero_renglon,
        numero_oc,
        cuit_oferente,
        cuit_proveedor,
    ) in registros_actuales:

        # Si esta fila ya tiene OC, se conserva.
        if numero_oc:
            continue

        if not numero_proceso or not numero_renglon:
            continue

        clave = (
            numero_proceso,
            numero_renglon,
        )

        empresas_adjudicadas = empresas_con_oc.get(
            clave,
            set()
        )

        if not empresas_adjudicadas:
            continue





        # Este Proceso + Renglón ya fue adjudicado.
        # Por lo tanto, esta fila queda identificada
        # como "Asig", independientemente de cuál
        # sea la empresa.

        ids_adjudicados_otra_empresa.append(
            registro_id
        )

        if not mostrar_adjudicados:
            ids_a_excluir.append(
                registro_id
            )



    # Mostrar siempre los registros adjudicados a otra empresa.
    # No excluir registros de la consulta.
    ids_a_excluir = []
    # ========================================================
    # COMPETIDORES POR PROCESO + RENGLÓN
    #
    # Funciona siempre, haya o no empresa seleccionada.
    # Busca en TODOS los lotes.
    # ========================================================

    competidores_por_clave = {}

    # --------------------------------------------------------
    # Obtener las combinaciones PROCESO + RENGLÓN
    # que aparecen en la consulta actual
    # --------------------------------------------------------

    claves_consulta = set(
        registros
        .exclude(numero_proceso__isnull=True)
        .exclude(numero_proceso="")
        .exclude(numero_renglon__isnull=True)
        .values_list(
            "numero_proceso",
            "numero_renglon",
        )
    )

    # --------------------------------------------------------
    # Buscar posibles competidores en TODOS los lotes
    # --------------------------------------------------------

    if claves_consulta:

        registros_competidores = (
            RegistroLicitacion.objects
            .filter(
                numero_proceso__in=[
                    clave[0]
                    for clave in claves_consulta
                ],
                numero_renglon__in=[
                    clave[1]
                    for clave in claves_consulta
                ],
            )
            .select_related(
                "empresa_oferente",
                "empresa_proveedor",
            )
        )

        # ----------------------------------------------------
        # Guardar empresas por combinación exacta
        # PROCESO + RENGLÓN
        # ----------------------------------------------------

        for otro in registros_competidores:

            clave = (
                otro.numero_proceso,
                otro.numero_renglon,
            )

            # Evitar mezclar proceso/renglón
            if clave not in claves_consulta:
                continue

            empresa = None

            if otro.empresa_oferente:

                empresa = otro.empresa_oferente

            elif otro.empresa_proveedor:

                empresa = otro.empresa_proveedor

            if not empresa:
                continue

            if clave not in competidores_por_clave:

                competidores_por_clave[clave] = {}


            if empresa.tipo_destinatario == "empresa":
                nombre_competidor = (
                    empresa.nombre_pila
                    or empresa.nombre
                    or ""
                )
            else:
                nombre_competidor = (
                    empresa.nombre
                    or ""
                )

            competidores_por_clave[
                clave
            ][empresa.cuit] = nombre_competidor
            
    # ========================================================
    # TOTALES POR EMPRESA
    #
    # Esta estructura alimenta la segunda pestaña de la pantalla.
    # Usa los mismos filtros de entrada que la tabla Resultados.
    #
    # REGLAS: 
    # 1) Para Empresa + Proceso + Renglón duplicados en distintos
    #    lotes, conservar el lote cuyo nombre tenga la fecha más
    #    reciente <= hoy.
    # 2) Si Proceso + Renglón tiene OC en cualquier registro de la
    #    base, excluir toda esa combinación del cálculo, sin
    #    importar qué proveedor tenga el OC.
    # 3) Luego agrupar por empresa y calcular procesos distintos,
    #    procesos/renglones y suma de ofertas.
    # ========================================================

    hoy = datetime.now().date()

    def fecha_del_nombre_lote(nombre):
        match = re.search(r"(\d{4}-\d{2}-\d{2})", nombre or "")
        if not match:
            return None
        try:
            return datetime.strptime(
                match.group(1),
                "%Y-%m-%d"
            ).date()
        except ValueError:
            return None

    # Todas las claves Proceso + Renglón que tienen OC en la base.
    # Esto es deliberadamente independiente de los filtros de Empresa,
    # Lote y OC: la regla dice que un OC de cualquier proveedor excluye
    # esa combinación completa del resumen.
    claves_con_oc = set(
        RegistroLicitacion.objects
        .exclude(numero_proceso__isnull=True)
        .exclude(numero_proceso="")
        .exclude(numero_renglon__isnull=True)
        .exclude(numero_renglon="")
        .exclude(
            Q(numero_oc__isnull=True) |
            Q(numero_oc="")
        )
        .values_list(
            "numero_proceso",
            "numero_renglon",
        )
    )

    # Evaluamos solamente los registros que resultan de los filtros
    # compartidos (Lote, Empresa, Estado y Proceso).
    # El filtro Con OC / Sin OC NO participa de esta pestaña.
    # La tabla Resultados no se modifica.
    registros_para_totales = list(
        registros_para_totales
        .select_related(
            "lote",
            "empresa_oferente",
            "empresa_proveedor",
        )
    )

    # Para cada Empresa + Proceso + Renglón elegimos un único registro.
    # La prioridad es la fecha del archivo más reciente <= hoy.
    # Si hay empate, usamos fecha de carga e id como desempate.
    seleccionados_por_clave = {}

    for registro in registros_para_totales:

        if not registro.numero_proceso or not registro.numero_renglon:
            continue

        clave_proceso = (
            registro.numero_proceso,
            registro.numero_renglon,
        )

        if clave_proceso in claves_con_oc:
            continue

        empresa_registro = None

        if registro.empresa_oferente:
            empresa_registro = registro.empresa_oferente
        elif registro.empresa_proveedor:
            empresa_registro = registro.empresa_proveedor

        if not empresa_registro:
            continue

        fecha_archivo = fecha_del_nombre_lote(
            registro.lote.nombre_archivo
        )

        # Si el archivo no tiene fecha YYYY-MM-DD o es futuro,
        # no puede ganar como lote vigente.
        fecha_valida = (
            fecha_archivo
            if fecha_archivo and fecha_archivo <= hoy
            else None
        )

        clave = (
            empresa_registro.cuit,
            registro.numero_proceso,
            registro.numero_renglon,
        )

        candidato = (
            fecha_valida,
            registro.lote.fecha_carga,
            registro.id,
            registro,
        )

        anterior = seleccionados_por_clave.get(clave)

        def peso(item):
            fecha, fecha_carga, registro_id, _ = item
            return (
                fecha is not None,
                fecha or datetime.min.date(),
                fecha_carga,
                registro_id,
            )

        if anterior is None or peso(candidato) > peso(anterior):
            seleccionados_por_clave[clave] = candidato

    # Agrupar los registros únicos por empresa.
    acumulado_empresas = {}

    for _, _, _, registro in seleccionados_por_clave.values():

        if registro.empresa_oferente:
            empresa_registro = registro.empresa_oferente
        elif registro.empresa_proveedor:
            empresa_registro = registro.empresa_proveedor
        else:
            continue

        cuit = empresa_registro.cuit

        if cuit not in acumulado_empresas:
            acumulado_empresas[cuit] = {
                "empresa": empresa_registro.nombre,
                "cuit": cuit,
                "procesos": set(),
                "procesos_renglones": 0,
                "total_ofertas": Decimal("0"),
            }

        datos = acumulado_empresas[cuit]

        datos["procesos"].add(
            registro.numero_proceso
        )

        datos["procesos_renglones"] += 1

        if registro.precio_total_oferta is not None:
            datos["total_ofertas"] += (
                registro.precio_total_oferta
            )

    totales_empresas = []

    for datos in acumulado_empresas.values():
        totales_empresas.append({
            "empresa": datos["empresa"],
            "cuit": datos["cuit"],
            "procesos_distintos": len(datos["procesos"]),
            "procesos_renglones": datos["procesos_renglones"],
            "total_ofertas": datos["total_ofertas"],
        })

    # Ordenar por Total Ofertas: mayor a menor.
    # En caso de empate, ordenar por empresa.
    totales_empresas.sort(
        key=lambda x: (-x["total_ofertas"], x["empresa"].lower())
    )

    total_procesos_distintos = len({
        (cuit, proceso)
        for cuit, datos in acumulado_empresas.items()
        for proceso in datos["procesos"]
    })

    total_procesos_renglones = sum(
        datos["procesos_renglones"]
        for datos in acumulado_empresas.values()
    )

    total_ofertas = sum(
        (
            datos["total_ofertas"]
            for datos in acumulado_empresas.values()
        ),
        Decimal("0")
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
    # ASIGNAR COMPETIDORES A CADA REGISTRO
    # ========================================================

    for registro in registros:

        clave = (
            registro.numero_proceso,
            registro.numero_renglon,
        )

        empresas_competidoras = (
            competidores_por_clave.get(
                clave,
                {}
            )
        )

        # ----------------------------------------------------
        # Identificar la empresa de esta fila
        # ----------------------------------------------------

        cuit_actual = None

        if registro.empresa_oferente:

            cuit_actual = (
                registro.empresa_oferente.cuit
            )

        elif registro.empresa_proveedor:

            cuit_actual = (
                registro.empresa_proveedor.cuit
            )

        # ----------------------------------------------------
        # Excluir a la propia empresa
        # ----------------------------------------------------

        nombres = [
            nombre
            for cuit, nombre
            in empresas_competidoras.items()
            if cuit != cuit_actual
        ]

        registro.competidores = " ; ".join(
            sorted(
                nombres,
                key=str.lower
            )
        )
            
    # ========================================================
    # ASIGNAR COMPETIDORES A CADA REGISTRO
    # ========================================================

    for registro in registros:

        clave = (
            registro.numero_proceso,
            registro.numero_renglon,
        )

        empresas_competidoras = (
            competidores_por_clave.get(
                clave,
                {}
            )
        )

        # ====================================================
        # EMPRESA DE LA FILA ACTUAL
        # ====================================================

        cuit_actual = None

        if registro.empresa_oferente:

            cuit_actual = (
                registro.empresa_oferente.cuit
            )

        elif registro.empresa_proveedor:

            cuit_actual = (
                registro.empresa_proveedor.cuit
            )

        # ====================================================
        # EXCLUIR LA EMPRESA DE LA FILA
        # ====================================================

        nombres = [
            nombre
            for cuit, nombre
            in empresas_competidoras.items()
            if cuit != cuit_actual
        ]

        registro.competidores = " ; ".join(
            sorted(
                nombres,
                key=str.lower
            )
        )
    # ========================================================
    # TOTAL DE OFERTAS DE LA TABLA RESULTADOS
    # ========================================================
    # Suma solamente los registros que actualmente se muestran
    # en la primera pestaana, respetando todos sus filtros.
    total_ofertas_resultados = sum(
        (
            registro.precio_total_oferta
            for registro in registros
            if registro.precio_total_oferta is not None
        ),
        Decimal("0")
    )

    # ========================================================
    # CONTEXTO
    # ========================================================

    contexto = {

        "lotes": lotes,
        "empresas": empresas,
        "archivos_iniciales": archivos_iniciales,
        "archivos_iniciales_seleccionados": archivos_iniciales_seleccionados,
        "todas_archivos": todas_archivos,
        "estados": estados,
        "registros": registros,
        "lotes_seleccionados": lotes_seleccionados,
        "empresa_seleccionada": empresa_seleccionada,
        "estado_seleccionado": estado,
        "oc_seleccionado": oc,
        "proceso_seleccionado": proceso,
        "mostrar_adjudicados": mostrar_adjudicados,
        "ids_adjudicados_otra_empresa": ids_adjudicados_otra_empresa,
        # Segunda pestaña: Totales por Empresa
        "totales_empresas": totales_empresas,
        "total_procesos_distintos": total_procesos_distintos,
        "total_procesos_renglones": total_procesos_renglones,
        "total_ofertas": total_ofertas,
        "total_ofertas_resultados": total_ofertas_resultados,
    }

    return render(
        request,
        "InfEje/consultar.html",
        contexto
    )


@require_GET
def api_vencimientos_carta(request):
    carta_id = request.GET.get("cartaId", "")

    if not carta_id:
        return JsonResponse({
            "ok": False,
            "error": "No se recibió el ID de la carta."
        }, status=400)

    try:
        definicion = CartaDefinicion.objects.get(pk=carta_id)
    except CartaDefinicion.DoesNotExist:
        return JsonResponse({
            "ok": False,
            "error": "No se encontró la definición de la carta."
        }, status=404)

    if not definicion.es_renovacion:
        return JsonResponse({
            "ok": False,
            "error": "La carta seleccionada no es una renovación."
        }, status=400)

    fecha_desde = definicion.fecha_desde
    fecha_hasta = definicion.fecha_hasta

    if not fecha_desde or not fecha_hasta:
        return JsonResponse({
            "ok": False,
            "error": "La carta no tiene configuradas las fechas Desde y Hasta."
        }, status=400)

    # 1. Buscar los contratos que vencen en el período.
    # 1. Buscar los contratos que vencen en el período.
    registros_vencimiento = (
        RegistroLicitacion.objects
        .filter(
            fecha_fin_contrato__range=(fecha_desde, fecha_hasta)
        )
        .exclude(numero_proceso__isnull=True)
        .exclude(numero_proceso="")
        .exclude(numero_renglon__isnull=True)
        .exclude(numero_renglon="")
        .exclude(fecha_apertura__isnull=True)
    )

    # 2. Obtener las claves exactas:
    #    Proceso + Renglón + Fecha de apertura.
    claves = set(
        registros_vencimiento.values_list(
            "numero_proceso",
            "numero_renglon",
            "fecha_apertura",
        ).distinct()
    )

    # 3. Recuperar todos los movimientos coincidentes,
    #    sin validar nuevamente la fecha de fin.
    from django.db.models import Q

    filtro = Q(pk__in=[])

    for proceso, renglon, apertura in claves:
        filtro |= Q(
            numero_proceso=proceso,
            numero_renglon=renglon,
            fecha_apertura=apertura,
        )

    registros = (
        RegistroLicitacion.objects
        .filter(filtro)
        .select_related(
            "empresa_oferente",
            "empresa_proveedor",
        )
    )

    # 4. Devolver únicamente empresas con movimientos.
    empresas = {}

    for registro in registros:
        for empresa in (
            registro.empresa_oferente,
            registro.empresa_proveedor,
        ):
            if not empresa or not empresa.email:
                continue

            empresas[empresa.id] = {
                "id": empresa.id,
                "nombre": empresa.nombre or "",
                "email": empresa.email or "",
            }

    print("====================================")
    print("CONTRATOS QUE VENCEN:", registros_vencimiento.count())
    print("MOVIMIENTOS COINCIDENTES:", registros.count())
    print("EMPRESAS DEVUELTAS:", len(empresas))
    print("====================================")

    return JsonResponse({
        "ok": True,
        "empresas": sorted(
            empresas.values(),
            key=lambda empresa: empresa["nombre"].lower()
        )
    })



@require_POST
def eliminar_lote(request, lote_id):
    # ========================================================
    # ELIMINAR LOTE Y SUS REGISTROS
    # ========================================================

    lote = get_object_or_404(Lote, id=lote_id)
    nombre = lote.nombre_archivo
    cantidad = lote.registros.count()

    # RegistroLicitacion.lote usa on_delete=models.CASCADE,
    # por lo que al eliminar el lote se eliminan sus registros.
    lote.delete()

    messages.success(
        request,
        f"Lote '{nombre}' eliminado correctamente. "
        f"Se eliminaron {cantidad} registros."
    )

    return redirect("consultar")


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

    # ========================================================
    # OBTENER TODOS LOS REGISTROS DE LA CONSULTA
    # (sin filtro de empresa)
    # para poder encontrar competidores
    # ========================================================

    registros_todos, _ = obtener_registros_filtrados(
        request,
        ignorar_empresa=True,
    )

    # ========================================================
    # ARMAR COMPETIDORES POR PROCESO + RENGLÓN
    # ========================================================

    competidores_por_clave = {}

    for otro in registros_todos:

        clave = (
            otro.numero_proceso,
            otro.numero_renglon,
        )

        empresa_otro = None

        if otro.empresa_oferente:
            empresa_otro = otro.empresa_oferente

        elif otro.empresa_proveedor:
            empresa_otro = otro.empresa_proveedor

        if not empresa_otro:
            continue

        if clave not in competidores_por_clave:
            competidores_por_clave[clave] = {}

        competidores_por_clave[clave][
            empresa_otro.cuit
        ] = empresa_otro.nombre

    # ========================================================
    # FILTRAR SOLAMENTE LA EMPRESA SELECCIONADA
    # ========================================================

    registros = registros_todos.filter(
        Q(cuit_oferente=empresa.cuit) |
        Q(cuit_proveedor=empresa.cuit)
    )

    # ========================================================
    # SACAR LA EMPRESA PROPIA DE LOS COMPETIDORES
    # ========================================================

    for clave, empresas_competidoras in competidores_por_clave.items():

        competidores_por_clave[clave] = [
            nombre
            for cuit, nombre
            in empresas_competidoras.items()
            if cuit != empresa.cuit
        ]

        competidores_por_clave[clave].sort(
            key=str.lower
        )

    # ========================================================
    # GENERAR PDF
    # ========================================================

    buffer = construir_pdf(
        registros,
        empresa=empresa,
        competidores_por_clave=competidores_por_clave,
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
    
def empresas(request):


    empresas = (
        Empresa.objects
        .all()
        .annotate(
            nombre_orden=Lower("nombre")
        )
        .order_by("nombre_orden")
    )
    total_empresas = empresas.count()

    total_email1 = (
        empresas
        .exclude(email__isnull=True)
        .exclude(email="")
        .count()
    )

    total_email2 = (
        empresas
        .exclude(email_2__isnull=True)
        .exclude(email_2="")
        .count()
    )

    total_email3 = (
        empresas
        .exclude(email_3__isnull=True)
        .exclude(email_3="")
        .count()
    )

    return render(
        request,
        "InfEje/empresas.html",
        {
            "empresas": empresas,
            "total_empresas": total_empresas,
            "total_email1": total_email1,
            "total_email2": total_email2,
            "total_email3": total_email3,
        }
    )



def exportar_empresas(request):
    """
    Exporta TODAS las empresas existentes en la tabla Empresa
    al Excel maestro PROVEEDORES_serv_limpieza.xlsx.

    El archivo se genera directamente desde la base de datos,
    por lo que no depende de los registros de licitaciones.
    """

    empresas = (
        Empresa.objects
        .all()
        .annotate(
            nombre_orden=Lower("nombre")
        )
        .order_by("nombre_orden")
    )

    workbook = Workbook()
    hoja = workbook.active
    hoja.title = "Empresas"

    # ========================================================
    # ENCABEZADOS
    # ========================================================

    hoja.append([
        "CUIT",
        "Nombre",
        "Nombre pila",
        "Archivo inicial",
        "mail1",
        "mail2",
        "mail3",
        "Telefono",
        "Provincia",
        "Tipo destinatario",
        "Comentarios",
    ])

    # ========================================================
    # EMPRESAS
    # ========================================================

    for empresa in empresas:
        hoja.append([
            empresa.cuit or "",
            empresa.nombre or "",
            empresa.nombre_pila or "",
            empresa.archivo_inicial or "",
            empresa.email or "",
            empresa.email_2 or "",
            empresa.email_3 or "",
            empresa.telefono or "",
            empresa.provincia or "",
            empresa.tipo_destinatario or "empresa",
            empresa.comentarios or "",
        ])

    # ========================================================
    # FORMATO BÁSICO
    # ========================================================

    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions

    anchos = {
        "A": 18,  # CUIT
        "B": 45,  # Nombre
        "C": 30,  # nombre pila
        "D": 35,  # Archivo inicial
        "E": 35,  # mail1
        "F": 35,  # mail2
        "G": 35,  # mail3
        "H": 22,  # Telefono
        "I": 22,  # Provincia
        "J": 20,  # Tipo destinatario
        "K": 50,  # Comentarios
    }

    for columna, ancho in anchos.items():
        hoja.column_dimensions[columna].width = ancho

    # ========================================================
    # RESPUESTA
    # ========================================================

    respuesta = HttpResponse(
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )

    respuesta["Content-Disposition"] = (
        'attachment; filename="PROVEEDORES_serv_limpieza.xlsx"'
    )

    workbook.save(respuesta)

    return respuesta


def importar_emails(request):

    if request.method == "POST":

        archivo = request.FILES.get(
            "archivo"
        )

        if not archivo:

            messages.error(
                request,
                "Debe seleccionar un archivo Excel."
            )

            return redirect(
                "importar_emails"
            )

        try:

            (
                actualizadas,
                nuevas,
                no_encontradas,
            ) = importar_emails_excel(
                archivo
            )

            mensaje = (
                f"Empresas actualizadas: "
                f"{actualizadas}. "
                f"Empresas nuevas: "
                f"{nuevas}."
            )

            if no_encontradas:

                mensaje += (
                    f" CUIT no encontrados: "
                    f"{len(no_encontradas)}."
                )

            messages.success(
                request,
                mensaje
            )

            return redirect(
                "empresas"
            )

        except Exception as error:

            messages.error(
                request,
                f"Error al importar mails: {error}"
            )

    return render(
        request,
        "InfEje/importar_emails.html"
    )

def guardar_emails(request, empresa_id):

    if request.method != "POST":
        return redirect("empresas")

    try:
        empresa = Empresa.objects.get(
            id=empresa_id
        )

    except Empresa.DoesNotExist:
        return redirect("empresas")

    # ========================================================
    # TELEFONO
    # ========================================================

    empresa.telefono = (
        request.POST.get("telefono") or None
    )

    # ========================================================
    # PROVINCIA
    # ========================================================

    empresa.provincia = (
        request.POST.get("provincia") or None
    )

    # ========================================================
    # EMAILS
    # ========================================================

    empresa.email = (
        request.POST.get("email") or None
    )

    empresa.email_2 = (
        request.POST.get("email_2") or None
    )

    empresa.email_3 = (
        request.POST.get("email_3") or None
    )

    # ========================================================
    # nombre_pila
    # ========================================================

    empresa.nombre_pila = (
        request.POST.get("nombre_pila") or None
    )


    # ========================================================
    # COMENTARIOS
    # ========================================================

    empresa.comentarios = (
        request.POST.get("comentarios") or None
    )
    # ========================================================

    empresa.datos_web = (
        int(request.POST.get("datos_web"))
        if request.POST.get("datos_web")
        else None
    )
    # ========================================================
    # archivo_inicial
    # ========================================================
    empresa.archivo_inicial = (
        request.POST.get("archivo_inicial") or None
    )


    empresa.tipo_destinatario = (
        "empresa"
        if request.POST.get("tipo_destinatario")
        else "persona"
    )

    empresa.genero = (
        request.POST.get("genero")
        or None
    )

    # ========================================================
    # GUARDAR
    # ========================================================

    empresa.save(
        update_fields=[

            "telefono",
            "provincia",
            "email",
            "email_2",
            "email_3",
            "comentarios",
            "datos_web",
            "archivo_inicial",
            "tipo_destinatario",
            "genero",
            "nombre_pila",
        ]
    )

    messages.success(
        request,
        f"Datos de {empresa.nombre} actualizados correctamente."
    )

    return redirect("empresas")
    


def armar_estrategia(compradores):
    """
    Arma el texto de estrategia según la cantidad
    de compradores únicos.
    """

    compradores = [
        comprador.strip()
        for comprador in compradores
        if comprador and comprador.strip()
    ]

    # Eliminar duplicados conservando el orden
    compradores_unicos = list(
        dict.fromkeys(compradores)
    )

    cantidad = len(compradores_unicos)

    if cantidad == 0:
        return ""

    if cantidad == 1:
        return compradores_unicos[0]

    if cantidad == 2:
        return (
            f"{compradores_unicos[0]} "
            f"y en {compradores_unicos[1]}"
        )

    if cantidad == 3:
        return (
            f"{compradores_unicos[0]}, "
            f"{compradores_unicos[1]} "
            f"y en {compradores_unicos[2]}"
        )

    return (
        f"{compradores_unicos[0]}, "
        f"{compradores_unicos[1]} "
        f"y en {compradores_unicos[2]} "
        f"entre otros"
    )

def exportar_csv_doppler(request):

    # ========================================================
    # TOMAR LOS REGISTROS SEGÚN LOS FILTROS DE LA PANTALLA
    # ========================================================

    registros, _ = obtener_registros_filtrados(request)

    # ========================================================
    # AGRUPAR POR CUIT
    # ========================================================

    empresas = {}

    for registro in registros:

        empresa = None

        if registro.empresa_oferente:
            empresa = registro.empresa_oferente

        elif registro.empresa_proveedor:
            empresa = registro.empresa_proveedor

        if not empresa:
            continue

        cuit = empresa.cuit

        if cuit not in empresas:

            empresas[cuit] = {
                "empresa": empresa,
                "compradores": [],
            }

        if registro.comprador:

            comprador = registro.comprador.strip()

            if (
                comprador
                and comprador
                not in empresas[cuit]["compradores"]
            ):

                empresas[cuit]["compradores"].append(
                    comprador
                )

    # ========================================================
    # CREAR CSV
    # ========================================================

    respuesta = HttpResponse(
        content_type="text/csv; charset=utf-8"
    )

    respuesta["Content-Disposition"] = (
        'attachment; filename="doppler_licitaciones.csv"'
    )

    # BOM para UTF-8
    respuesta.write("\ufeff")

    escritor = csv.writer(
        respuesta,
        delimiter=";",
        lineterminator="\n",
    )

    # ========================================================
    # ENCABEZADO
    # ========================================================

    escritor.writerow([
        "EMAIL",
        "NOMBRE",
        "APELLIDO",
        "APORTE",
        "LOCALIDAD",
        "CLIENTE",
        "estrategia",
        "FECHA_CREACION",
    ])

    fecha_creacion = datetime.now().strftime(
        "%d/%m/%Y"
    )

    # ========================================================
    # EMPRESAS
    # ========================================================

    for datos in empresas.values():

        empresa = datos["empresa"]

        nombre_empresa = (
            empresa.nombre or ""
        ).strip()

        estrategia = armar_estrategia(
            datos["compradores"]
        )

        cliente = nombre_empresa

        # ----------------------------------------------------
        # PERSONA / EMPRESA
        # ----------------------------------------------------
        cuit = str(empresa.cuit or "").strip()

        es_persona = cuit.startswith(
            ("20", "23", "24", "27")
        )

        if es_persona:

            partes = nombre_empresa.split()

            if len(partes) >= 2:

                apellido = partes[-1]

                nombre = " ".join(
                    partes[:-1]
                )

            else:

                nombre = nombre_empresa
                apellido = ""

            aporte = ""

        else:

            nombre = nombre_empresa
            apellido = ""
            aporte = ""

        # ----------------------------------------------------
        # EMAILS
        # ----------------------------------------------------

        emails = [
            empresa.email,
            empresa.email_2,
            empresa.email_3,
        ]

        emails_validos = []

        for email in emails:

            if not email:
                continue

            email = str(email).strip()

            if not email:
                continue

            if email.lower() in [
                e.lower()
                for e in emails_validos
            ]:
                continue

            emails_validos.append(email)

        # ----------------------------------------------------
        # UNA FILA POR EMAIL
        # ----------------------------------------------------

        for email in emails_validos:

            escritor.writerow([
                email,
                nombre,
                apellido,
                aporte,
                "",
                cliente,
                estrategia,
                fecha_creacion,
            ])

    return respuesta

def guardar_ficha_empresa(request, empresa_id):

    if request.method != "POST":
        return redirect("consultar")

    try:
        empresa = Empresa.objects.get(
            id=empresa_id
        )

    except Empresa.DoesNotExist:
        return redirect("consultar")

    empresa.telefono = (
        request.POST.get("telefono") or None
    )

    empresa.email = (
        request.POST.get("email") or None
    )

    empresa.email_2 = (
        request.POST.get("email_2") or None
    )

    empresa.email_3 = (
        request.POST.get("email_3") or None
    )
    empresa.nombre_pila = (
        request.POST.get("nombre_pila") or None
    )
    empresa.comentarios = (
        request.POST.get("comentarios") or None
    )
    empresa.novedades = (
        request.POST.get("novedades") or None
    )
    empresa.tipo_destinatario = (
        "empresa" if request.POST.get("tipo_destinatario") else "persona"
    )
    empresa.save(
        update_fields=[
            "telefono",
            "email",
            "email_2",
            "email_3",
            "comentarios",
            "novedades",
            "tipo_destinatario",
            "nombre_pila",
        ]
    )



    if request.headers.get("X-Requested-With") == "XMLHttpRequest":
        return JsonResponse({
            "ok": True,
            "mensaje": "Información guardada correctamente."
        })

    messages.success(
        request,
        f"Información de {empresa.nombre} actualizada correctamente."
    )

    return redirect(
        f"{reverse('consultar')}?{request.GET.urlencode()}"
    )





def exportar_word_mails(request):

    carta_id = request.GET.get("carta")



    definicion_carta = CartaDefinicion.objects.filter(
        id=carta_id,
        activa=True,
    ).first()

    if not definicion_carta:
        return HttpResponse(
            "La carta seleccionada no existe o está inactiva.",
            status=400,
        )

    carta = str(definicion_carta.codigo)



    print("ID RECIBIDO:", carta_id)
    print("ID DEFINICIÓN:", definicion_carta.id)
    print("CÓDIGO DE LA CARTA:", definicion_carta.codigo)
    print("NOMBRE DE LA CARTA:", definicion_carta.nombre)

    # ========================================================
    # TOMAR LOS REGISTROS SEGÚN LOS FILTROS DE LA PANTALLA
    # ========================================================

    registros, _ = obtener_registros_filtrados(request)

    registros = (
        registros
        .select_related(
            "empresa_oferente",
            "empresa_proveedor",
            "lote",
        )
    )



    # ========================================================
    # RENOVACIONES: BUSCAR CONTRATOS Y SUS PARTICIPANTES
    # ========================================================

    if definicion_carta and definicion_carta.es_renovacion:

        fecha_desde = definicion_carta.fecha_desde
        fecha_hasta = definicion_carta.fecha_hasta

        # Validar las fechas configuradas en Editar Carta
        if not fecha_desde or not fecha_hasta:
            return HttpResponse(
                "Para generar la carta de renovación, "
                "tenés que configurar las fechas Desde y Hasta.",
                status=400,
            )

        if fecha_desde > fecha_hasta:
            return HttpResponse(
                "La fecha Desde no puede ser posterior "
                "a la fecha Hasta.",
                status=400,
            )

        # 1. Encontrar los registros cuyo contrato vence
        #    dentro del período configurado.
        registros_vencimiento = (
            RegistroLicitacion.objects
            .filter(
                fecha_fin_contrato__range=(
                    fecha_desde,
                    fecha_hasta,
                )
            )
            .exclude(numero_proceso__isnull=True)
            .exclude(numero_proceso="")
            .exclude(numero_renglon__isnull=True)
            .exclude(numero_renglon="")
            .exclude(fecha_apertura__isnull=True)
        )
        print("====================================")
        print("CARTA:", definicion_carta.nombre)
        print("ES RENOVACIÓN:", definicion_carta.es_renovacion)
        print("DESDE:", fecha_desde)
        print("HASTA:", fecha_hasta)
        print("CONTRATOS ENCONTRADOS:", registros_vencimiento.count())
        print("====================================")
        # 2. Obtener las combinaciones exactas:
        #    Proceso + Renglón + Fecha de apertura.
        claves_renovacion = set(
            registros_vencimiento.values_list(
                "numero_proceso",
                "numero_renglon",
                "fecha_apertura",
            ).distinct()
        )

        # 3. Buscar todos los movimientos coincidentes,
        #    SIN filtrar por fecha_fin_contrato.
        filtro_claves = Q(pk__in=[])

        for proceso, renglon, apertura in claves_renovacion:
            filtro_claves |= Q(
                numero_proceso=proceso,
                numero_renglon=renglon,
                fecha_apertura=apertura,
            )

        registros = (
            RegistroLicitacion.objects
            .filter(filtro_claves)
            .select_related(
                "empresa_oferente",
                "empresa_proveedor",
                "lote",
            )
        )

        # ====================================================
        # GUARDAR EL ADJUDICATARIO DE CADA PROCESO
        # ANTES DE FILTRAR LAS EMPRESAS SELECCIONADAS
        # ====================================================

        adjudicatarios_por_clave = {}

        registros_con_oc = registros.exclude(
            Q(numero_oc__isnull=True) | Q(numero_oc="")
        )

        for registro in registros_con_oc:

            clave = (
                registro.numero_proceso,
                registro.numero_renglon,
                registro.fecha_apertura,
            )

            empresa_ganadora = (
                registro.empresa_proveedor
                or registro.empresa_oferente
            )

            if empresa_ganadora:
                adjudicatarios_por_clave.setdefault(
                    clave, set()
                ).add(empresa_ganadora.nombre or "")





        print("MOVIMIENTOS RECUPERADOS:", registros.count())


        # ====================================================
        # IDENTIFICAR ADJUDICATARIOS DE CADA PROCESO
        # ====================================================

        adjudicatarios_por_clave = {}

        registros_con_oc = (
            registros
            .exclude(Q(numero_oc__isnull=True) | Q(numero_oc=""))
            .select_related(
                "empresa_proveedor",
                "empresa_oferente",
            )
        )

        for registro in registros_con_oc:

            clave = (
                registro.numero_proceso,
                registro.numero_renglon,
                registro.fecha_apertura,
            )

            empresa_ganadora = (
                registro.empresa_proveedor
                or registro.empresa_oferente
            )

            if empresa_ganadora:

                adjudicatarios_por_clave.setdefault(
                    clave, set()
                ).add(empresa_ganadora.nombre or "")



        print(
            "EMPRESAS CON MOVIMIENTOS:",
            registros.values(
                "empresa_oferente_id",
                "empresa_proveedor_id",
            ).distinct().count()
        )

    # ========================================================
    # EXCLUIR EN WORD PROCESOS + RENGLONES
    # YA ADJUDICADOS
    #
    # Si el Proceso + Renglón tiene OC:
    # - la fila que tiene OC se conserva
    # - cualquier otra fila sin OC se excluye
    #   aunque sea de la misma empresa adjudicataria.
    # ========================================================

    mostrar_adjudicados = True

    if not mostrar_adjudicados and carta != "4":
        claves_con_oc = set(
            RegistroLicitacion.objects
            .exclude(
                Q(numero_oc__isnull=True) |
                Q(numero_oc="")
            )
            .exclude(numero_proceso__isnull=True)
            .exclude(numero_proceso="")
            .exclude(numero_renglon__isnull=True)
            .exclude(numero_renglon="")
            .values_list(
                "numero_proceso",
                "numero_renglon",
            )
        )

        ids_a_excluir_word = []

        for registro in registros:

            if registro.numero_oc:
                continue

            if (
                not registro.numero_proceso
                or not registro.numero_renglon
            ):
                continue

            clave = (
                registro.numero_proceso,
                registro.numero_renglon,
            )

            if clave in claves_con_oc:
                ids_a_excluir_word.append(
                    registro.id
                )

        if ids_a_excluir_word:
            registros = registros.exclude(
                id__in=ids_a_excluir_word
            )

            
    # ========================================================
    # EMPRESAS SELECCIONADAS PARA LA PLANILLA DE ENVÍOS
    # ========================================================

    empresa_ids = request.GET.getlist(
        "empresa_ids"
    )

    prima_minima_pesos = Decimal(
        request.GET.get(
            "prima_minima_pesos",
            "18000"
        )
    )

    tc = Decimal(
        request.GET.get(
            "tc",
            "1510"
        )
    )


    if empresa_ids:

        # Solamente empresas que todavía tienen
        # al menos un registro después de aplicar
        # los filtros de Word.
        empresas_con_registros_ids = set()

        for registro in registros:


            if registro.empresa_oferente_id:
                empresas_con_registros_ids.add(registro.empresa_oferente_id)

            if registro.empresa_proveedor_id:
                empresas_con_registros_ids.add(registro.empresa_proveedor_id)
                
        empresas_para_registro = (
            Empresa.objects
            .filter(
                id__in=empresa_ids
            )
            .filter(
                id__in=empresas_con_registros_ids
            )
            .order_by("nombre")
        )

    else:
        empresas_para_registro = None












    # ========================================================
    # FILTRO ADICIONAL: EMPRESAS SELECCIONADAS PARA WORD
    # ========================================================

    empresa_ids = request.GET.getlist("empresa_ids")

    if empresa_ids:

        registros = registros.filter(
            Q(
                empresa_oferente_id__in=empresa_ids
            )
            |
            Q(
                empresa_proveedor_id__in=empresa_ids
            )
        )



    # ====================================================
    # ASIGNAR NOMBRE DEL ADJUDICATARIO A CADA REGISTRO
    # ====================================================

    if definicion_carta and definicion_carta.es_renovacion:

        for registro in registros:

            clave = (
                registro.numero_proceso,
                registro.numero_renglon,
                registro.fecha_apertura,
            )

            registro.nombre_adjudicatario = ", ".join(
                sorted(
                    adjudicatarios_por_clave.get(clave, set())
                )
            )


        
    # ========================================================
    # COMPETIDORES POR PROCESO + RENGLÓN
    # ========================================================

    competidores_por_clave = {}
    # Asociar el adjudicatario a cada movimiento,
    # aunque la empresa ganadora no esté seleccionada.

    if definicion_carta.es_renovacion:

        for registro in registros:

            clave = (
                registro.numero_proceso,
                registro.numero_renglon,
                registro.fecha_apertura,
            )

            registro.nombre_adjudicatario = ", ".join(
                sorted(
                    adjudicatarios_por_clave.get(clave, set())
                )
            )



    claves_consulta = set(
        registros
        .exclude(numero_proceso__isnull=True)
        .exclude(numero_proceso="")
        .exclude(numero_renglon__isnull=True)
        .values_list(
            "numero_proceso",
            "numero_renglon",
        )
    )

    if claves_consulta:

        registros_competidores = (
            RegistroLicitacion.objects
            .filter(
                numero_proceso__in=[
                    clave[0]
                    for clave in claves_consulta
                ],
                numero_renglon__in=[
                    clave[1]
                    for clave in claves_consulta
                ],
            )
            .select_related(
                "empresa_oferente",
                "empresa_proveedor",
            )
        )

        for otro in registros_competidores:

            clave = (
                otro.numero_proceso,
                otro.numero_renglon,
            )

            if clave not in claves_consulta:
                continue

            empresa = None

            if otro.empresa_oferente:
                empresa = otro.empresa_oferente

            elif otro.empresa_proveedor:
                empresa = otro.empresa_proveedor

            if not empresa:
                continue

            if clave not in competidores_por_clave:
                competidores_por_clave[clave] = {}


            if empresa.tipo_destinatario == "empresa":
                nombre_competidor = (
                    empresa.nombre_pila                    
                )
            else:
                nombre_competidor = empresa.nombre

            competidores_por_clave[
                clave
            ][empresa.cuit] = nombre_competidor
    # ========================================================
    # ASIGNAR COMPETIDORES A CADA REGISTRO
    # ========================================================

    for registro in registros:

        clave = (
            registro.numero_proceso,
            registro.numero_renglon,
        )

        empresas_competidoras = (
            competidores_por_clave.get(
                clave,
                {}
            )
        )

        # ----------------------------------------------------
        # Identificar la empresa de esta fila
        # ----------------------------------------------------

        cuit_actual = None

        if registro.empresa_oferente:
            cuit_actual = (
                registro.empresa_oferente.cuit
            )

        elif registro.empresa_proveedor:
            cuit_actual = (
                registro.empresa_proveedor.cuit
            )

        # ----------------------------------------------------
        # Excluir la propia empresa
        # ----------------------------------------------------

        nombres = [
            nombre
            for cuit, nombre
            in empresas_competidoras.items()
            if cuit != cuit_actual
        ]


        registro.competidores = " ; ".join(
            sorted(
                (nombre for nombre in nombres if nombre),
                key=str.lower
            )
        )



    # ========================================================
    # AGRUPAR POR EMPRESA
    # ========================================================

    empresas = {}


    for registro in registros:

        empresas_del_registro = []

        if registro.empresa_oferente:
            empresas_del_registro.append(
                registro.empresa_oferente
            )

        # En las renovaciones hay que asociar el movimiento
        # tanto al oferente como al proveedor.
        if (
            registro.empresa_proveedor
            and (
                (definicion_carta and definicion_carta.es_renovacion)
                or carta == "4"
                or not registro.empresa_oferente
            )
        ):
            empresas_del_registro.append(
                registro.empresa_proveedor
            )

        for empresa in empresas_del_registro:

            # Si se seleccionaron empresas para el Word,
            # incluir únicamente esas empresas.
            if empresa_ids and str(empresa.id) not in empresa_ids:
                continue

            if empresa.id not in empresas:
                empresas[empresa.id] = {
                    "empresa": empresa,
                    "registros": [],
                }

            empresas[empresa.id]["registros"].append(
                registro
            )


    # ========================================================
    # ORDENAR EMPRESAS
    # ========================================================

    empresas_ordenadas = sorted(
        empresas.values(),
        key=lambda x: (
            x["empresa"].nombre or ""
        ).lower()
    )

    empresas_con_registros = [
        (
            datos["empresa"],
            datos["registros"],
        )
        for datos in empresas_ordenadas
    ]

    # ========================================================
    # DATOS DE COTIZACIÓN
    # ========================================================

    prima_minima_pesos = Decimal(
        request.GET.get(
            "prima_minima_pesos",
            "18000"
        )
    )

    tc = Decimal(
        request.GET.get(
            "tc",
            "1510"
        )
    )


    # ========================================================
    # GENERAR WORD
    # ========================================================

    # La definición ya fue obtenida al comienzo de la función.
    # "carta" ya contiene su código.

    if definicion_carta:
        imprimir_movimientos = definicion_carta.incluir_movimientos
        incluir_cotizacion = definicion_carta.incluir_cotizacion
    else:
        imprimir_movimientos = (
            request.GET.get("imprimir_movimientos", "0") == "1"
        )
        incluir_cotizacion = False


    fecha_desde = request.GET.get("fechaDesdeVencimiento", "")
    fecha_hasta = request.GET.get("fechaHastaVencimiento", "")

    buffer = generar_word_mails(
        empresas_con_registros,
        empresas_para_registro=empresas_para_registro,
        fecha_envio=date.today().strftime("%d/%m/%Y"),
        prima_minima_pesos=prima_minima_pesos,
        tc=tc,
        carta=carta,
        imprimir_movimientos=imprimir_movimientos,
        incluir_cotizacion=incluir_cotizacion,
        fecha_desde=fecha_desde,
        fecha_hasta=fecha_hasta,
    )



    respuesta = HttpResponse(
        buffer.getvalue(),
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )

    respuesta["Content-Disposition"] = (
        'attachment; '
        'filename="mails_licitaciones.docx"'
    )

    return respuesta


from django.db.models import Count
from django.shortcuts import render


def seguimiento_cartas(request):

    empresas = (
        Empresa.objects
        .all()
        .order_by("nombre")
    )

    # Se muestran TODOS los eventos del flujo.
    seguimientos = (
        SeguimientoCartasFlow.objects
        .select_related("empresa", "carta")
        .order_by(
            "empresa__nombre",
            "fecha_envio",
            "id",
        )
    )

    # Total de movimientos registrados.
    total_movimientos = seguimientos.count()

    # Cantidad de movimientos agrupados por la carta que se muestra
    # en la tabla. Los registros sin carta se agrupan por evento.
    resumen_eventos = (
        seguimientos
        .values("carta__nombre")
        .annotate(cantidad=Count("id"))
        .order_by("carta__nombre")
    )

    # Cartas configuradas y activas, ordenadas por nombre.



    cartas_activas = (
        CartaDefinicion.objects
        .filter(activa=True)
        .prefetch_related("condiciones")
        .order_by("nombre")
    )

    for carta in cartas_activas:
        print(f"\nCARTA: {carta.nombre} | CÓDIGO: {carta.codigo} | ID: {carta.id}")

        for condicion in carta.condiciones.all():
            print(
                f"    {condicion.tipo} -> "
                f"{condicion.codigo_evento}"
            )

    return render(
        request,
        "InfEje/seguimiento_cartas.html",
        {
            "empresas": empresas,
            "seguimientos": seguimientos,
            "total_movimientos": total_movimientos,
            "resumen_eventos": resumen_eventos,
            "cartas_activas": cartas_activas,
        }
    )


def exportar_excel_seguimiento_cartas(request):

    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from django.http import HttpResponse


    # =========================================================
    # TODOS LOS EVENTOS Y CARTAS CONFIGURADAS
    # =========================================================

    seguimientos = (
        SeguimientoCartasFlow.objects
        .select_related("empresa", "carta")
        .order_by("empresa__nombre", "id")
    )

    cartas_definidas = list(
        CartaDefinicion.objects
        .prefetch_related("condiciones")
        .order_by("nombre")
    )

    # Relacionar los códigos de evento con las cartas configuradas.
    cartas_por_evento = {}

    for carta_def in cartas_definidas:
        for condicion in carta_def.condiciones.all():
            if condicion.tipo == "incluir":
                cartas_por_evento.setdefault(
                    condicion.codigo_evento,
                    carta_def
                )

    def obtener_carta_registro(registro):
        # Si el movimiento tiene una carta asociada, usarla.
        if registro.carta_id and registro.carta:
            return registro.carta

        # Para movimientos anteriores sin carta asociada,
        # buscar la definición mediante sus condiciones.
        return cartas_por_evento.get(registro.evento)

    # =========================================================
    # AGRUPAR POR PROVEEDOR
    # =========================================================

    proveedores = {}

    for seguimiento in seguimientos:
        proveedores.setdefault(
            seguimiento.empresa_id, []
        ).append(seguimiento)

    # =========================================================
    # CLASIFICAR PROVEEDORES
    # =========================================================
    # =========================================================
    # CLASIFICAR PROVEEDORES SEGÚN LA ÚLTIMA CARTA ENVIADA
    # =========================================================


    # ============================================================
    # CLASIFICAR PROVEEDORES SEGÚN LA ÚLTIMA CARTA
    # ============================================================

    destinos_excel = {
        (carta.agrupacion_destino or "demas").strip().lower()
        for carta in cartas_definidas
    }
    destinos_excel.add("demas")

    grupos_por_destino = {
        destino: [] for destino in destinos_excel
    }

    for empresa_id, registros in proveedores.items():

        # Conservar todos los movimientos del proveedor
        registros = sorted(
            registros,
            key=lambda r: (
                r.fecha_envio is None,
                r.fecha_envio,
                r.id
            )
        )

        # Buscar el último movimiento enviado que tenga una carta identificada
        registros_con_fecha = [
            registro
            for registro in registros
            if registro.fecha_envio is not None
            and obtener_carta_registro(registro) is not None
        ]

        if registros_con_fecha:
            ultimo_registro = registros_con_fecha[-1]
            ultima_carta = obtener_carta_registro(ultimo_registro)

            destino = (
                ultima_carta.agrupacion_destino or "demas"
            ).strip().lower()
        else:
            destino = "demas"

        # Si el destino no existe, enviar a Demás
        if destino not in grupos_por_destino:
            destino = "demas"

        # Se guarda el proveedor con TODOS sus movimientos
        grupos_por_destino[destino].append(registros)
 

    # =========================================================
    # CREAR EXCEL
    # =========================================================



    wb = Workbook()

    # Destinos que se utilizan en el Excel.
    destinos_excel = {
        (carta.agrupacion_destino or "demas").strip().lower()
        for carta in cartas_definidas
    }

    destinos_excel.add("demas")

    # Relacionar cada destino con su pestaña.
    hojas_excel = {}

    hoja_inicial = wb.active
    wb.remove(hoja_inicial)

    for destino in sorted(destinos_excel):

        if destino == "demas":
            nombre_hoja = "Demás"
        else:
            carta_destino = next(
                (
                    carta
                    for carta in cartas_definidas
                    if carta.codigo == destino
                ),
                None
            )

            if carta_destino:
                nombre_hoja = carta_destino.nombre[:31]
            else:
                nombre_hoja = destino[:31]

        ws = wb.create_sheet(nombre_hoja)
        hojas_excel[destino] = ws


    # =========================================================
    # ESTILOS
    # =========================================================

    fuente_titulo = Font(
        bold=True
    )

    borde_separador = Border(
        bottom=Side(
            style="thin"
        )
    )

    # =========================================================
    # FUNCIÓN PARA ESCRIBIR CADA PESTAÑA
    # =========================================================


    def escribir_pestana(ws, grupos, titulo, resumen=False):
        # Título en A1
        celda_titulo = ws["A1"]
        celda_titulo.value = titulo
        celda_titulo.font = Font(bold=True)


        # Borde superior e inferior desde A hasta F
        borde_titulo = Border(
            top=Side(style="medium"),
            bottom=Side(style="medium")
        )

        for columna in range(1, 6):
            ws.cell(
                row=1,
                column=columna
            ).border = borde_titulo


        # Resumen solamente para la pestaña Demás


        # Resumen solamente para la pestaña Demás
        if resumen:

            total_proveedores = len(grupos)

            ws["A2"] = "Proveedores"
            ws["B2"] = total_proveedores

            fila_resumen = 3

            # Mostrar únicamente las cartas que tienen movimientos,
            # con sus nombres reales y cantidades.
            for carta_def in cartas_definidas:

                cantidad = sum(
                    1
                    for registros in grupos
                    for registro in registros
                    if (
                        (carta_registro := obtener_carta_registro(registro))
                        is not None
                        and carta_registro.pk == carta_def.pk
                    )
                )

                if cantidad:
                    ws.cell(
                        row=fila_resumen,
                        column=1,
                        value=carta_def.nombre
                    )
                    ws.cell(
                        row=fila_resumen,
                        column=2,
                        value=cantidad
                    )
                    fila_resumen += 1

            for fila_resumen_estilo in range(2, fila_resumen):
                for columna in (1, 2):
                    celda = ws.cell(
                        row=fila_resumen_estilo,
                        column=columna
                    )
                    celda.font = Font(bold=True)
                    celda.alignment = Alignment(horizontal="left")

            # Dejar una fila vacía antes de los encabezados.
            fila = fila_resumen + 1

        else:
            fila = 3

        encabezados = [
            "Empresa",
            "Evento",
            "Fecha envío",
            "Fecha respuesta",
            "Teléfono/Mails/Provincia",
        ]


        borde_encabezado = Border(
            top=Side(style="medium"),
            bottom=Side(style="medium")
        )

        for columna, encabezado in enumerate(encabezados, start=1):

            celda = ws.cell(
                row=fila,
                column=columna,
                value=encabezado
            )

            celda.font = fuente_titulo
            celda.border = borde_encabezado


        fila += 1

        # -----------------------------------------
        # PROVEEDORES
        # -----------------------------------------

        for registros in grupos:

            # -----------------------------------------
            # IMPRIMIR TODAS LAS CARTAS DEL PROVEEDOR
            # -----------------------------------------

            for registro in registros:

                ws.cell(
                    row=fila,
                    column=1,
                    value=registro.empresa.nombre
                )

                ws.cell(
                    row=fila,
                    column=2,
                    value=(
                        registro.carta.nombre
                        if registro.carta
                        else registro.evento
                    )
                )


                fecha_envio = registro.fecha_envio
                fecha_respuesta = registro.fecha_respuesta

                celda_fecha_envio = ws.cell(
                    row=fila,
                    column=3,
                    value=fecha_envio
                )

                celda_fecha_respuesta = ws.cell(
                    row=fila,
                    column=4,
                    value=fecha_respuesta
                )

                if fecha_envio is not None:
                    celda_fecha_envio.number_format = "dd/mm/yyyy"

                if fecha_respuesta is not None:
                    celda_fecha_respuesta.number_format = "dd/mm/yyyy"

                fila += 1

            # -----------------------------------------
            # CONTACTOS: UNA SOLA FILA DESPUÉS DE TODAS LAS CARTAS
            # -----------------------------------------

            empresa = registros[0].empresa

            contactos = [
                empresa.telefono,
                empresa.email,
                empresa.email_2,
                empresa.email_3,
                empresa.provincia,
            ]


            # Datos de WEB: agregar solo si tiene valor
            if empresa.datos_web is not None:
                contactos.append(
                    empresa.get_datos_web_display()
                )

            contactos = [
                str(dato).strip()
                for dato in contactos
                if dato and str(dato).strip()
            ]

            celda_contactos = ws.cell(
                row=fila,
                column=5,
                value="\n".join(contactos)
            )

            celda_contactos.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )

            # Datos de WEB en rojo
            if empresa.datos_web is not None:
                celda_contactos.font = Font(color="FF0000")

            fila += 1


            # -----------------------------------------
            # LÍNEA ENTRE PROVEEDORES
            # -----------------------------------------

            for columna in range(1, 6):
                ws.cell(
                    row=fila,
                    column=columna
                ).border = borde_separador

            fila += 1

        # -----------------------------------------
        # ANCHO DE COLUMNAS
        # -----------------------------------------


        anchos = {
            "A": 35,
            "B": 18,
            "C": 15,
            "D": 18,
            "E": 45,
        }

        for columna, ancho in anchos.items():

            ws.column_dimensions[
                columna
            ].width = ancho

    # =========================================================
    # ESCRIBIR LAS PESTAÑAS DINÁMICAS
    # =========================================================

    fecha_informe = date.today().strftime("%d/%m/%Y")


    for destino, ws in hojas_excel.items():

        grupos = grupos_por_destino.get(destino, [])

        if destino == "demas":
            titulo = (
                f"LICITACIONES - PROVEEDORES - al {fecha_informe}"
            )
            resumen = True
        else:

            carta_destino = next(
                (
                    carta for carta in cartas_definidas
                    if (carta.codigo or "").strip().lower() == destino
                ),
                None
            )
            nombre_destino = (
                carta_destino.nombre
                if carta_destino
                else destino
            )

            titulo = (
                f"LICITACIONES - PROVEEDORES - "
                f"{nombre_destino} al {fecha_informe}"
            )
            resumen = False

        escribir_pestana(
            ws,
            grupos,
            titulo,
            resumen=resumen
        )


    # =========================================================
    # RESPUESTA
    # =========================================================

    respuesta = HttpResponse(
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )



    respuesta[
        "Content-Disposition"
    ] = (
        'attachment; '
        f'filename="Licitaciones-Informe {date.today().strftime("%Y-%m-%d")}.xlsx"'
    )


    wb.save(
        respuesta
    )

    return respuesta



@require_POST
def aplicar_cartas_enviadas(request):

    try:
        print("========== APLICAR CARTAS ENVIADAS ==========")
        datos = json.loads(request.body)
        envios = datos.get("envios", [])




        carta_valor = datos.get("carta")

        carta_obj = None

        if carta_valor:
            try:
                carta_obj = CartaDefinicion.objects.filter(
                    id=int(carta_valor)
                ).first()
            except (ValueError, TypeError):
                carta_obj = None


            # Si no existe ese código, el selector puede haber enviado el ID.
            if carta_obj is None and str(carta_valor).isdigit():
                carta_obj = CartaDefinicion.objects.filter(
                    id=int(carta_valor)
                ).first()


                
        if not envios:
            return JsonResponse(
                {
                    "ok": False,
                    "error": "No hay eventos para aplicar."
                },
                status=400
            )

        cantidad = 0

        for envio in envios:

            empresa_id = envio.get("empresa_id")
            evento = envio.get("evento")
            fecha_texto = envio.get("fecha")

            if not empresa_id or not evento or not fecha_texto:
                continue

            try:
                fecha = datetime.strptime(
                    fecha_texto,
                    "%Y-%m-%d"
                ).date()

            except (ValueError, TypeError):
                continue

            empresa = Empresa.objects.filter(
                id=empresa_id
            ).first()

            if not empresa:
                continue
            print("CARTA RECIBIDA:", carta_valor)
            print("CARTA ENCONTRADA:", carta_obj)
            print("APLICANDO EVENTO:", empresa.id, evento, fecha)
            SeguimientoCartasFlow.objects.create(
                empresa=empresa,
                carta=carta_obj,
                evento=evento,
                fecha_envio=fecha,
            )
            cantidad += 1

        if cantidad == 0:
            return JsonResponse(
                {
                    "ok": False,
                    "error": "No se pudo registrar ningún evento."
                },
                status=400
            )

        return JsonResponse(
            {
                "ok": True,
                "cantidad": cantidad
            }
        )

    except Exception as e:

        print(
            "ERROR APLICAR EVENTO CARTA:",
            e
        )

        return JsonResponse(
            {
                "ok": False,
                "error": str(e)
            },
            status=500
        )







@require_POST
def guardar_seguimiento_carta(request):
    """
    Crea nuevos eventos en el historial del flujo.

    Cada cambio de etapa agrega un nuevo registro.
    Nunca se modifica el evento anterior.
    """

    try:
        datos = json.loads(request.body)

        seguimientos = datos.get("seguimientos", [])

        if not seguimientos:
            return JsonResponse(
                {
                    "ok": False,
                    "error": "No hay eventos para guardar."
                },
                status=400
            )

        cantidad = 0

        for dato in seguimientos:
            empresa_id = dato.get("empresa_id")
            evento = dato.get("evento")

            if not empresa_id or not evento:
                continue

            empresa = Empresa.objects.filter(
                id=empresa_id
            ).first()

            if not empresa:
                continue

            fecha_envio = None
            fecha_texto = dato.get("fecha")

            if fecha_texto:
                try:
                    fecha_envio = datetime.strptime(
                        fecha_texto,
                        "%Y-%m-%d"
                    ).date()
                except (ValueError, TypeError):
                    fecha_envio = None
            print("APLICANDO EVENTO:", empresa.id, evento, fecha)
            SeguimientoCartasFlow.objects.create(
                empresa=empresa,
                evento=evento,
                fecha_envio=fecha_envio,
                acciones=dato.get("acciones") or None,
            )

            cantidad += 1

        if cantidad == 0:
            return JsonResponse(
                {
                    "ok": False,
                    "error": "No se pudo guardar ningún evento."
                },
                status=400
            )

        return JsonResponse(
            {
                "ok": True,
                "cantidad": cantidad
            }
        )

    except Exception as e:
        print(
            "ERROR GUARDAR SEGUIMIENTO:",
            e
        )

        return JsonResponse(
            {
                "ok": False,
                "error": str(e)
            },
            status=500
        )


@require_POST
def eliminar_seguimiento_carta(request):

    try:

        datos = json.loads(request.body)
        ids = datos.get("ids", [])

        if not ids:
            return JsonResponse({
                "ok": False,
                "error": "No hay cartas seleccionadas."
            }, status=400)

        SeguimientoCartasFlow.objects.filter(
            id__in=ids
        ).delete()

        return JsonResponse({
            "ok": True
        })

    except Exception as e:

        return JsonResponse(
            {
                "ok": False,
                "error": str(e)
            },
            status=500
        )

@xframe_options_sameorigin

def editar_carta(request):

    codigo_carta = (
        request.POST.get("carta")
        or request.GET.get("carta")
        or "1"
    )


    definicion_id = (
        request.POST.get("definicion_id")
        or request.GET.get("definicion_id")
    )

    if definicion_id in ("None", "null", ""):
        definicion_id = None
    # ------------------------------------------------------------
    # TIPO DE CARTA: EMPRESA, MUJER U HOMBRE
    # ------------------------------------------------------------

    destinatario_vista = (
        request.POST.get("destinatario_vista")
        or request.GET.get("destinatario_vista")
        or "empresa"
    )

    if destinatario_vista == "mujer":
        tipo_destinatario = "persona"
        genero = "mujer"
    elif destinatario_vista == "hombre":
        tipo_destinatario = "persona"
        genero = "hombre"
    else:
        destinatario_vista = "empresa"
        tipo_destinatario = "empresa"
        genero = None

    # ------------------------------------------------------------
    # EMPRESA PARA VISTA PREVIA
    # ------------------------------------------------------------

    empresas = Empresa.objects.all().order_by("nombre")

    empresa_id = (
        request.GET.get("empresa_id")
        or request.POST.get("empresa_id")
    )

    empresa_seleccionada = None

    if empresa_id:
        empresa_seleccionada = Empresa.objects.filter(
            pk=empresa_id
        ).first()

    # ------------------------------------------------------------
    # DEFINICIÓN SELECCIONADA
    # ------------------------------------------------------------

    definicion = None

    if codigo_carta != "__nueva__":
        definicion = CartaDefinicion.objects.filter(
            codigo=codigo_carta
        ).first()

        if definicion is None and definicion_id:
            definicion = CartaDefinicion.objects.filter(
                pk=definicion_id,
                codigo=codigo_carta
            ).first()


    # Si una definición anterior no tenía código, asignarle uno.
    if definicion and not definicion.codigo:
        definicion.codigo = f"CARTA_{definicion.pk}"
        definicion.save(update_fields=["codigo"])

    # ------------------------------------------------------------
    # CARGAR LA PLANTILLA PARA LA VISTA ACTUAL
    # ------------------------------------------------------------

    codigo_plantilla = (
        str(definicion.codigo)
        if definicion and definicion.codigo
        else str(codigo_carta)
    )

    plantilla = CartaPlantilla.objects.filter(
        carta=codigo_plantilla,
        tipo_destinatario=tipo_destinatario,
        genero=genero,
    ).first()

    # Si todavía no existe, preparar una instancia vacía
    # sin crear un registro en la base de datos.
    if plantilla is None:
        plantilla = CartaPlantilla(
            carta=codigo_plantilla,
            tipo_destinatario=tipo_destinatario,
            genero=genero,
        )




    # ------------------------------------------------------------
    # ELIMINAR CARTA PERSONALIZADA
    # ------------------------------------------------------------

    if (
        request.method == "POST"
        and request.POST.get("accion") == "eliminar"
    ):
        try:
            definicion_eliminar = definicion

            # Buscar por ID si no se encontró por código.
            if definicion_eliminar is None and definicion_id:
                definicion_eliminar = CartaDefinicion.objects.filter(
                    pk=definicion_id
                ).first()

            codigos_protegidos = {
                "1", "2", "3",
                *[f"R{i}" for i in range(1, 13)],
                "Fin", "Click",
            }

            if definicion_eliminar is None:
                messages.error(
                    request,
                    "No se encontró la carta que querés eliminar."
                )

            elif (
                not definicion_eliminar.codigo
                or definicion_eliminar.codigo in codigos_protegidos
            ):
                messages.error(
                    request,
                    "No se puede eliminar una carta predefinida."
                )

            else:
                with transaction.atomic():
                    codigo_eliminar = definicion_eliminar.codigo

                    # Las plantillas guardan el código como texto.
                    CartaPlantilla.objects.filter(
                        carta=str(codigo_eliminar)
                    ).delete()

                    # Las condiciones se eliminan por CASCADE.
                    definicion_eliminar.delete()

                messages.success(
                    request,
                    "La carta y sus plantillas se eliminaron correctamente."
                )

        except Exception as error:
            messages.error(
                request,
                f"No se pudo eliminar la carta: {error}"
            )

        return redirect(
            f"/editar-carta/?carta=1"
            f"&empresa_id={empresa_id or ''}"
            f"&destinatario_vista={destinatario_vista}"
        )




    # ------------------------------------------------------------
    # GUARDAR
    # ------------------------------------------------------------

    if request.method == "POST":

        nombre = (request.POST.get("nombre") or "").strip()
        codigo = (request.POST.get("codigo") or "").strip()

        fecha_desde_texto = request.POST.get("fecha_desde") or ""
        fecha_hasta_texto = request.POST.get("fecha_hasta") or ""

        fecha_desde = (
            parse_date(fecha_desde_texto)
            if fecha_desde_texto else None
        )

        fecha_hasta = (
            parse_date(fecha_hasta_texto)
            if fecha_hasta_texto else None
        )

        es_nueva = codigo_carta == "__nueva__"

        if es_nueva and not nombre:
            messages.error(
                request,
                "Ingresá el nombre de la nueva carta."
            )

        elif (
            (fecha_desde_texto and not fecha_desde)
            or (fecha_hasta_texto and not fecha_hasta)
        ):
            messages.error(
                request,
                "Ingresá fechas válidas."
            )

        elif (
            fecha_desde
            and fecha_hasta
            and fecha_desde > fecha_hasta
        ):
            messages.error(
                request,
                "La fecha Desde no puede ser posterior a Hasta."
            )

        else:
            try:
                with transaction.atomic():

                    # Crear o actualizar la definición.
                    if definicion is None:
                        definicion = CartaDefinicion()
                        # Si se está guardando una carta predefinida (1, 2, 3, R1...)
                        # usar ese mismo código. Si se genera CARTA_<id>, al volver
                        # a buscarla el selector no encuentra la configuración guardada.
                        if not es_nueva and codigo_carta:
                            definicion.codigo = str(codigo_carta)
                            nombres_base = {
                                "1": "Carta 1. Presentación",
                                "2": "Carta 2. Simulación",
                                "3": "Carta 3. Invitación",
                                "Fin": "Fin",
                                "Click": "Click",
                            }
                            if str(codigo_carta).startswith("R") and str(codigo_carta)[1:].isdigit():
                                definicion.nombre = f"Renov {int(str(codigo_carta)[1:]):02d}"
                            else:
                                definicion.nombre = nombres_base.get(str(codigo_carta), str(codigo_carta))

                    if nombre:
                        definicion.nombre = nombre

                    if codigo:
                        definicion.codigo = codigo

                    definicion.tipo_destinatario = tipo_destinatario
                    definicion.genero = genero

                    definicion.incluir_movimientos = (
                        request.POST.get("incluir_movimientos") == "on"
                    )

                    definicion.incluir_cotizacion = (
                        request.POST.get("incluir_cotizacion") == "on"
                    )

                    definicion.requiere_email = (
                        request.POST.get("requiere_email") == "on"
                    )

                    if request.method == "POST" and definicion:
                        definicion.todos_proveedores = (
                            request.POST.get("todos_proveedores") == "on"
                        )

                    definicion.es_renovacion = (
                        request.POST.get("es_renovacion") == "on"
                    )


                    definicion.agrupacion_destino = (
                        request.POST.get("agrupacion_destino", "demas")
                    )

                    definicion.fecha_desde = fecha_desde
                    definicion.fecha_hasta = fecha_hasta

                    # En una carta nueva queda activa por defecto.
                    definicion.activa = (
                        request.POST.get("activa") == "on"
                        if "activa" in request.POST
                        else es_nueva
                    )






                    definicion.save()

                    if not definicion.codigo:
                        definicion.codigo = (
                            f"CARTA_{definicion.pk}"
                        )
                        definicion.save(update_fields=["codigo"])

                    # ------------------------------------------------
                    # GUARDAR CONDICIONES
                    # Solo si el formulario contiene esos campos.
                    # ------------------------------------------------

                    if "condiciones_presentes" in request.POST:

                        incluir = request.POST.getlist(
                            "condicion_incluir"
                        )

                        excluir = request.POST.getlist(
                            "condicion_excluir"
                        )

                        definicion.condiciones.all().delete()

                        condiciones = set()

                        for evento in incluir:
                            evento = evento.strip()
                            if evento:
                                condiciones.add(("incluir", evento))

                        for evento in excluir:
                            evento = evento.strip()
                            if evento:
                                condiciones.add(("excluir", evento))

                        for tipo_condicion, evento in sorted(condiciones):
                            CartaCondicion.objects.create(
                                carta=definicion,
                                tipo=tipo_condicion,
                                codigo_evento=evento,
                            )

                    # ------------------------------------------------
                    # GUARDAR ASUNTO Y CONTENIDO DE LA PLANTILLA
                    # ------------------------------------------------

                    plantilla, _ = CartaPlantilla.objects.get_or_create(
                        carta=str(definicion.codigo),
                        tipo_destinatario=tipo_destinatario,
                        genero=genero,
                    )

                    plantilla.asunto = (
                        request.POST.get("asunto") or ""
                    )

                    plantilla.contenido_html = (
                        request.POST.get("contenido_html") or ""
                    )

                    plantilla.save()



                return redirect(
                    f"/editar-carta/?carta={definicion.codigo}"
                    f"&definicion_id={definicion.pk}"
                    f"&empresa_id={empresa_id or ''}"
                    f"&destinatario_vista={destinatario_vista}"
                )

            except Exception as error:
                messages.error(
                    request,
                    f"No se pudo guardar la carta: {error}"
                )

    # ------------------------------------------------------------
    # LISTAS DE LOS COMBOS Y CARTAS
    # ------------------------------------------------------------


    cartas_definidas = sorted(
        list(
            CartaDefinicion.objects
            .exclude(codigo__isnull=True)
            .exclude(codigo="")
            .values("codigo", "nombre")
        ),
        key=lambda carta: carta["nombre"].casefold()
    )

    # NOMBRE COMPLETO DE LA CARTA SELECCIONADA
    nombre_carta = next(
        (
            item["nombre"]
            for item in cartas_definidas
            if item["codigo"] == codigo_carta
        ),
        codigo_carta
    )

    # ------------------------------------------------------------
    # NOMBRE COMPLETO DE LA CARTA SELECCIONADA
    # ------------------------------------------------------------

    nombre_carta = next(
        (
            item["nombre"]
            for item in cartas_definidas
            if item["codigo"] == codigo_carta
        ),
        codigo_carta
    )

    # ------------------------------------------------------------
    # OPCIONES PARA LOS COMBOS INCLUIR Y EXCLUIR
    # ------------------------------------------------------------

    opciones_condicion = list(
        CartaDefinicion.objects
        .filter(activa=True)
        .order_by("nombre")
        .values_list("id", "nombre")
    )

    opciones_condicion = [
        {
            "codigo": str(carta_id),
            "nombre": nombre,
        }
        for carta_id, nombre in opciones_condicion
    ]

    incluir_seleccionados = set()
    excluir_seleccionados = set()

    if definicion:
        incluir_seleccionados = set(
            definicion.condiciones.filter(tipo="incluir")
            .values_list("codigo_evento", flat=True)
        )

        excluir_seleccionados = set(
            definicion.condiciones.filter(tipo="excluir")
            .values_list("codigo_evento", flat=True)
        )

    # ------------------------------------------------------------
    # MOSTRAR PANTALLA
    # ------------------------------------------------------------
    
    return render(
        request,
        "InfEje/editar_carta.html",
        {
            "plantilla": plantilla,
            "empresas": empresas,
            "empresa_seleccionada": empresa_seleccionada,
            "definicion": definicion,


            "puede_eliminar": (
                bool(definicion and definicion.codigo)
                and definicion.codigo not in {
                    "1", "2", "3",
                    *[f"R{i}" for i in range(1, 13)],
                    "Fin", "Click",
                }
            ),



            "cartas_definidas": cartas_definidas,
            "condiciones": (
                definicion.condiciones.all()
                if definicion else []
            ),
            "opciones_condicion": opciones_condicion,
            "incluir_seleccionados": incluir_seleccionados,
            "excluir_seleccionados": excluir_seleccionados,
            "es_carta_nueva": codigo_carta == "__nueva__",
            "destinatario_vista": destinatario_vista,
            "codigo_carta": codigo_carta,
            "nombre_carta": nombre_carta,



            
        }
    )




def obtener_plantilla_carta(request, carta, tipo_destinatario):

    destinatario_vista = (
        request.GET.get("destinatario_vista")
        or request.POST.get("destinatario_vista")
    )

    if destinatario_vista == "mujer":
        tipo_destinatario = "persona"
        genero = "mujer"
    elif destinatario_vista == "hombre":
        tipo_destinatario = "persona"
        genero = "hombre"
    elif destinatario_vista == "empresa":
        tipo_destinatario = "empresa"
        genero = None
    else:
        genero = request.GET.get("genero") or None

    print("====================================")
    print("CARGANDO PLANTILLA")
    print("Carta:", carta)
    print("Tipo:", tipo_destinatario)
    print("Genero:", genero)
    print("====================================")

    plantilla = CartaPlantilla.objects.filter(
        carta=str(carta),
        tipo_destinatario=tipo_destinatario,
        genero=genero,
    ).first()

    if not plantilla:
        return JsonResponse(
            {"error": "No existe una plantilla para esa carta y destinatario."},
            status=404,
        )

    return JsonResponse({
        "asunto": plantilla.asunto or "",
        "contenido_html": plantilla.contenido_html or "",
    })


@require_POST
def guardar_plantilla_carta(request):

    try:

        datos = json.loads(request.body)

        carta = str(datos.get("carta"))

        destinatario_vista = datos.get("destinatario_vista")

        if destinatario_vista == "mujer":
            tipo_destinatario = "persona"
            genero = "mujer"
        elif destinatario_vista == "hombre":
            tipo_destinatario = "persona"
            genero = "hombre"
        else:
            tipo_destinatario = "empresa"
            genero = None
        asunto = datos.get("asunto") or ""
        contenido_html = datos.get("contenido_html") or ""

        print("====================================")
        print("GUARDANDO PLANTILLA")
        print("Carta:", carta)
        print("Tipo:", tipo_destinatario)
        print("Genero:", genero)
        print("Asunto:", asunto)
        print("Contenido:", contenido_html[:100])
        print("====================================")

        plantilla, creada = CartaPlantilla.objects.get_or_create(
            carta=carta,
            tipo_destinatario=tipo_destinatario,
            genero=genero
        )

        plantilla.asunto = asunto
        plantilla.contenido_html = contenido_html
        plantilla.save()

        return JsonResponse({
            "ok": True,
            "creada": creada
        })

    except Exception as e:

        print("ERROR GUARDAR PLANTILLA:", e)

        return JsonResponse({
            "ok": False,
            "error": str(e)
        }, status=500)

# ============================================================
# CONFIGURACIÓN DE CARTAS DESDE INFORMESLICITACIONES
# ============================================================

def configurar_cartas(request):

    carta_id = (
        request.POST.get("carta_id")
        or request.GET.get("id")
    )

    carta = None

    if carta_id:
        carta = get_object_or_404(
            CartaDefinicion,
            pk=carta_id
        )

    # ========================================================
    # GUARDAR CARTA Y CONDICIONES
    # ========================================================

    if request.method == "POST":

        nombre = (
            request.POST.get("nombre") or ""
        ).strip()

        codigo = (
            request.POST.get("codigo") or ""
        ).strip() or None

        if not nombre:
            messages.error(
                request,
                "El nombre de la carta es obligatorio."
            )

        else:

            fecha_desde_texto = (
                request.POST.get("fecha_desde") or ""
            )

            fecha_hasta_texto = (
                request.POST.get("fecha_hasta") or ""
            )

            fecha_desde = (
                parse_date(fecha_desde_texto)
                if fecha_desde_texto else None
            )

            fecha_hasta = (
                parse_date(fecha_hasta_texto)
                if fecha_hasta_texto else None
            )

            if (
                fecha_desde_texto and not fecha_desde
            ) or (
                fecha_hasta_texto and not fecha_hasta
            ):
                messages.error(
                    request,
                    "Ingresá fechas válidas."
                )

            elif (
                fecha_desde
                and fecha_hasta
                and fecha_desde > fecha_hasta
            ):
                messages.error(
                    request,
                    "La fecha Desde no puede ser posterior a Hasta."
                )

            else:

                try:

                    with transaction.atomic():

                        if carta is None:
                            carta = CartaDefinicion()

                        carta.codigo = codigo
                        carta.nombre = nombre

                        carta.tipo_destinatario = (
                            request.POST.get(
                                "tipo_destinatario",
                                "empresa"
                            )
                        )

                        carta.genero = (
                            request.POST.get("genero") or None
                        )

                        carta.incluir_movimientos = (
                            request.POST.get(
                                "incluir_movimientos"
                            ) == "on"
                        )

                        carta.incluir_cotizacion = (
                            request.POST.get(
                                "incluir_cotizacion"
                            ) == "on"
                        )

                        carta.requiere_email = (
                            request.POST.get(
                                "requiere_email"
                            ) == "on"
                        )

                        carta.es_renovacion = (
                            request.POST.get(
                                "es_renovacion"
                            ) == "on"
                        )

                        carta.fecha_desde = fecha_desde
                        carta.fecha_hasta = fecha_hasta

                        carta.activa = (
                            request.POST.get("activa") == "on"
                        )

                        carta.save()

                        # ------------------------------------
                        # CONDICIONES DE INCLUSION Y EXCLUSION
                        # ------------------------------------

                        codigos_incluir = (
                            request.POST.getlist(
                                "condicion_incluir"
                            )
                        )

                        codigos_excluir = (
                            request.POST.getlist(
                                "condicion_excluir"
                            )
                        )

                        # Reemplazamos las condiciones anteriores
                        # por las seleccionadas en los combos.
                        carta.condiciones.all().delete()

                        condiciones_guardadas = set()

                        for tipo, codigos in (
                            ("incluir", codigos_incluir),
                            ("excluir", codigos_excluir),
                        ):

                            for codigo_evento in codigos:

                                codigo_evento = (
                                    codigo_evento or ""
                                ).strip()

                                if not codigo_evento:
                                    continue

                                clave = (
                                    tipo,
                                    codigo_evento
                                )

                                # Evitar condiciones duplicadas.
                                if clave in condiciones_guardadas:
                                    continue

                                CartaCondicion.objects.create(
                                    carta=carta,
                                    tipo=tipo,
                                    codigo_evento=codigo_evento
                                )

                                condiciones_guardadas.add(clave)

                    messages.success(
                        request,
                        "La carta y sus condiciones "
                        "se guardaron correctamente."
                    )

                    return redirect(
                        "configurar_cartas"
                    )

                except Exception as error:

                    messages.error(
                        request,
                        f"No se pudo guardar la carta: {error}"
                    )

    # ========================================================
    # LISTADO Y FORMULARIO
    # ========================================================

    cartas = CartaDefinicion.objects.all()

    condiciones = (
        carta.condiciones.all()
        if carta else []
    )

    # Códigos seleccionados para mostrar las casillas marcadas.
    incluir_seleccionados = set()
    excluir_seleccionados = set()

    for condicion in condiciones:
        if condicion.tipo == "incluir":
            incluir_seleccionados.add(
                condicion.codigo_evento
            )
        elif condicion.tipo == "excluir":
            excluir_seleccionados.add(
                condicion.codigo_evento
            )

    return render(
        request,
        "InfEje/configurar_cartas.html",
        {
            "cartas": cartas,
            "carta": carta,
            "condiciones": condiciones,
            "incluir_seleccionados": incluir_seleccionados,
            "excluir_seleccionados": excluir_seleccionados,
        }
    )


    
from django.http import JsonResponse

def api_cartas_activas(request):
    cartas = (
        CartaDefinicion.objects
        .filter(activa=True)
        .prefetch_related("condiciones")
        .order_by("nombre")
    )
    for carta in cartas:
        print(
            "ID:", carta.id,
            "| Código:", carta.codigo,
            "| Nombre:", carta.nombre
        )
    return JsonResponse({
        "cartas": [
            {
                "valor": str(carta.id),
                "nombre": carta.nombre,
                "todos_proveedores": carta.todos_proveedores,
                "condiciones": [
                    {
                        "tipo": condicion.tipo,
                        "codigo_evento": condicion.codigo_evento,
                    }
                    for condicion in carta.condiciones.all()
                ],
            }
            for carta in cartas
        ]
    })