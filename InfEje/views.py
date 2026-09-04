from django.shortcuts import render
from django.db.models import (
    Q,
    Case,
    When,
    Value,
    IntegerField,
    CharField,
)
from django.db.models.functions import Coalesce, Cast, Lower

from .models import (
    Lote,
    Empresa,
    RegistroLicitacion
)
import csv
import re
from openpyxl import Workbook
from decimal import Decimal
from datetime import datetime
from django.contrib import messages
from django.shortcuts import render, redirect, get_object_or_404
from django.urls import reverse
from .importadores import importar_excel
from django.http import HttpResponse, JsonResponse
from django.views.decorators.http import require_POST

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

        cuit_empresa = (
            cuit_oferente
            or cuit_proveedor
        )

        if not cuit_empresa:
            continue

        # Si otra empresa tiene la OC,
        # esta fila queda identificada como "Asig".
        if cuit_empresa not in empresas_adjudicadas:

            ids_adjudicados_otra_empresa.append(
                registro_id
            )

            if not mostrar_adjudicados:
                ids_a_excluir.append(
                    registro_id
                )


    # Con "No" ocultamos esas filas.
    # Con "Sí" quedan visibles y se marcarán como "Asig".
    if ids_a_excluir:
        registros = registros.exclude(
            id__in=ids_a_excluir
        )
                
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

            competidores_por_clave[
                clave
            ][empresa.cuit] = empresa.nombre

            
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
        "Archivo inicial",
        "mail1",
        "mail2",
        "mail3",
        "Telefono",
        "Provincia",
        "Comentarios",
    ])

    # ========================================================
    # EMPRESAS
    # ========================================================

    for empresa in empresas:
        hoja.append([
            empresa.cuit or "",
            empresa.nombre or "",
            empresa.archivo_inicial or "",
            empresa.email or "",
            empresa.email_2 or "",
            empresa.email_3 or "",
            empresa.telefono or "",
            empresa.provincia or "",
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
        "C": 35,  # Archivo inicial
        "D": 35,  # mail1
        "E": 35,  # mail2
        "F": 35,  # mail3
        "G": 22,  # Telefono
        "H": 22,  # Provincia
        "I": 50,  # Comentarios
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
                no_encontradas,
            ) = importar_emails_excel(
                archivo
            )

            mensaje = (
                f"Empresas actualizadas: "
                f"{actualizadas}."
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
    # COMENTARIOS
    # ========================================================

    empresa.comentarios = (
        request.POST.get("comentarios") or None
    )

    # ========================================================
    # archivo_inicial
    # ========================================================
    empresa.archivo_inicial = (
        request.POST.get("archivo_inicial") or None
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
            "archivo_inicial",
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

    empresa.comentarios = (
        request.POST.get("comentarios") or None
    )
    empresa.novedades = (
        request.POST.get("novedades") or None
    )
    empresa.save(
        update_fields=[
            "telefono",
            "email",
            "email_2",
            "email_3",
            "comentarios",
            "novedades",
        ]
    )

    messages.success(
        request,
        f"Información de {empresa.nombre} actualizada correctamente."
    )

    return redirect(
        f"{reverse('consultar')}?{request.GET.urlencode()}"
    )
