import pandas as pd

from django.db import transaction

from .models import (
    Lote,
    Empresa,
    RegistroLicitacion,
)


# ============================================================
# MAPEO DE COLUMNAS DEL EXCEL
# ============================================================

MAPEO_COLUMNAS = {
    "Unidad de Negocios": "unidad_negocios",

    "Regla": "segmento",
    "Segmento": "segmento",

    "Número de proceso": "numero_proceso",
    "Portal de compra": "portal_compra",
    "Comprador": "comprador",
    "Estado": "estado",

    "Número de renglón": "numero_renglon",
    "Descripción del renglón": "descripcion_renglon",
    "Cantidad": "cantidad",
    "Fecha de apertura AAAA-MM-DD": "fecha_apertura",

    "Oferente": "oferente",
    "Identificador del oferente": "cuit_oferente",

    "Alternativa": "alternativa",
    "Especificación técnica del oferente": "especificacion_tecnica",
    "Cantidad ofertada": "cantidad_ofertada",

    "Precio unitario de oferta": "precio_unitario_oferta",
    "Precio total de oferta": "precio_total_oferta",
    "Moneda seleccionada": "moneda_oferta",

    "Número de OC": "numero_oc",
    "Fecha de inicio de contrato": "fecha_inicio_contrato",
    "Fecha de fin de contrato": "fecha_fin_contrato",
    "Duración del contrato (días)": "duracion_contrato",
    "Tipo de contrato": "tipo_contrato",

    "Proveedor": "proveedor",
    "Identificador del proveedor": "cuit_proveedor",

    "Descripción del contrato": "descripcion_contrato",
    "Cantidad comprada": "cantidad_comprada",

    "Precio unitario de compra": "precio_unitario_compra",
    "Precio total de compra": "precio_total_compra",
    "Moneda seleccionada.1": "moneda_compra",

    "URL de la fuente de datos": "url",
}


# ============================================================
# VALORES VACÍOS
# ============================================================

def valor_o_none(valor):

    if pd.isna(valor):
        return None

    if isinstance(valor, str):

        valor = valor.strip()

        if not valor:
            return None

    return valor


# ============================================================
# LIMPIAR CUIT
# ============================================================

def limpiar_cuit(valor):

    valor = valor_o_none(valor)

    if valor is None:
        return None

    try:
        return str(int(float(valor)))

    except (ValueError, TypeError):

        return str(valor).strip()


def limpiar_telefono(valor):
    """
    Convierte teléfonos provenientes de Excel
    evitando que queden con .0 cuando Excel
    los entrega como número.
    """

    valor = valor_o_none(valor)

    if valor is None:
        return None

    if isinstance(valor, float):
        if valor.is_integer():
            return str(int(valor))

    if isinstance(valor, int):
        return str(valor)

    return str(valor).strip()


# ============================================================
# LIMPIAR FECHA
# ============================================================
def limpiar_fecha(valor):
    """
    Convierte fechas provenientes de Excel a date.
    Si está vacío o no se puede convertir, devuelve None.
    """

    valor = valor_o_none(valor)

    if valor is None:
        return None

    try:
        fecha = pd.to_datetime(
            valor,
            errors="coerce"
        )

        if pd.isna(fecha):
            return None

        return fecha.date()

    except (ValueError, TypeError):
        return None

# ============================================================
# EMPRESA
# ============================================================
def obtener_empresa(cuit, nombre, archivo_inicial=None):
    cuit = limpiar_cuit(cuit)

    if not cuit:
        return None, False

    nombre = valor_o_none(nombre)

    if not nombre:
        nombre = "Empresa sin nombre"

    empresa, creada = Empresa.objects.get_or_create(
        cuit=cuit,
        defaults={
            "nombre": nombre,
            "archivo_inicial": archivo_inicial,
        }
    )

    # Si la empresa ya existía y todavía no tiene
    # archivo_inicial, guardamos el archivo que estamos importando.
    if not creada and not empresa.archivo_inicial and archivo_inicial:
        empresa.archivo_inicial = archivo_inicial
        empresa.save(update_fields=["archivo_inicial"])

    return empresa, creada


# ============================================================
# IMPORTAR EXCEL
# ============================================================

@transaction.atomic
def importar_excel(archivo, nombre_archivo):

    # --------------------------------------------------------
    # Evitar duplicar un lote
    # --------------------------------------------------------

    if Lote.objects.filter(
        nombre_archivo=nombre_archivo
    ).exists():

        raise ValueError(
            f"El archivo '{nombre_archivo}' ya fue importado."
        )

    # --------------------------------------------------------
    # Leer cuarta pestaña
    # --------------------------------------------------------

    df = pd.read_excel(
        archivo,
        sheet_name=3,
        header=3
    )

    # --------------------------------------------------------
    # Limpiar nombres de columnas
    # --------------------------------------------------------

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    # --------------------------------------------------------
    # Crear lote
    # --------------------------------------------------------

    lote = Lote.objects.create(
        nombre_archivo=nombre_archivo,
        rubro="Servicios de limpieza",
    )

    registros = []

    empresas_creadas = 0

    # --------------------------------------------------------
    # Procesar filas
    # --------------------------------------------------------

    for _, fila in df.iterrows():

        datos = {}

        for columna_excel, campo_modelo in MAPEO_COLUMNAS.items():

            if columna_excel not in df.columns:
                continue

            valor = fila[columna_excel]



            if campo_modelo in (
                "cuit_oferente",
                "cuit_proveedor",
            ):

                valor = limpiar_cuit(valor)

            elif campo_modelo in (
                "fecha_apertura",
                "fecha_inicio_contrato",
                "fecha_fin_contrato",
            ):

                valor = limpiar_fecha(valor)

            elif campo_modelo == "duracion_contrato":

                valor = valor_o_none(valor)

                if valor is not None:
                    try:
                        valor = int(float(valor))
                    except (ValueError, TypeError):
                        valor = None

            else:

                valor = valor_o_none(valor)





            datos[campo_modelo] = valor

        # ----------------------------------------------------
        # EMPRESA OFERENTE
        # ----------------------------------------------------

        empresa_oferente, creada = obtener_empresa(
            datos.get("cuit_oferente"),
            datos.get("oferente"),
            nombre_archivo,
        )

        if empresa_oferente:

            datos["empresa_oferente"] = empresa_oferente

            if creada:
                empresas_creadas += 1

        # ----------------------------------------------------
        # EMPRESA PROVEEDOR
        # ----------------------------------------------------

        empresa_proveedor, creada = obtener_empresa(
            datos.get("cuit_proveedor"),
            datos.get("proveedor"),
            nombre_archivo,
        )

        if empresa_proveedor:

            datos["empresa_proveedor"] = empresa_proveedor

            if creada:
                empresas_creadas += 1

        # ----------------------------------------------------
        # LOTE
        # ----------------------------------------------------

        datos["lote"] = lote

        registros.append(
            RegistroLicitacion(**datos)
        )

    # --------------------------------------------------------
    # Guardar registros
    # --------------------------------------------------------

    RegistroLicitacion.objects.bulk_create(
        registros,
        batch_size=1000
    )

    return (
        lote,
        len(registros),
        empresas_creadas,
    )


# ============================================================
# DIAGNÓSTICO
# ============================================================

def diagnosticar_columnas(
    ruta_archivo,
    nombre_hoja
):

    df = pd.read_excel(
        ruta_archivo,
        sheet_name=nombre_hoja
    )

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    columnas_excel = set(df.columns)

    columnas_esperadas = set(
        MAPEO_COLUMNAS.keys()
    )

    encontradas = sorted(
        columnas_excel.intersection(
            columnas_esperadas
        )
    )

    faltantes = sorted(
        columnas_esperadas.difference(
            columnas_excel
        )
    )

    print("\n========================================")
    print("ARCHIVO:", ruta_archivo)
    print("HOJA:", nombre_hoja)
    print("========================================")

    print("\nCOLUMNAS ENCONTRADAS:")

    for columna in encontradas:
        print("  OK  ", columna)

    print("\nCOLUMNAS NO PRESENTES:")

    for columna in faltantes:
        print("  --- ", columna)

    print("\nCOLUMNAS NO MAPEADAS:")

    for columna in sorted(
        columnas_excel - columnas_esperadas
    ):
        print("  ??? ", columna)

    print("\n========================================\n")

    return df.columns.tolist()

def importar_emails_excel(archivo):
    """
    Actualiza la información de Empresa utilizando exclusivamente el CUIT.

    El Excel debe tener:
        CUIT
        Nombre
        Archivo inicial
        mail1
        mail2
        mail3
        Telefono
        Provincia
        Comentarios

    Lo que viene del Excel reemplaza lo que existe actualmente
    en la base de datos.

    Novedades NO se modifica.
    """

    df = pd.read_excel(
        archivo
    )

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    columnas_obligatorias = {
        "CUIT",
        "Nombre",
        "Archivo inicial",
        "mail1",
        "mail2",
        "mail3",
        "Telefono",
        "Provincia",
        "Comentarios",
    }

    faltantes = (
        columnas_obligatorias
        - set(df.columns)
    )

    if faltantes:
        raise ValueError(
            "Faltan columnas obligatorias: "
            + ", ".join(sorted(faltantes))
        )

    actualizadas = 0
    no_encontradas = []

    for _, fila in df.iterrows():

        # =====================================================
        # CUIT
        # =====================================================

        cuit = limpiar_cuit(
            fila["CUIT"]
        )

        if not cuit:
            continue

        # =====================================================
        # BUSCAR EMPRESA POR CUIT
        # =====================================================

        try:

            empresa = Empresa.objects.get(
                cuit=cuit
            )

        except Empresa.DoesNotExist:

            no_encontradas.append(
                cuit
            )

            continue

        # =====================================================
        # NOMBRE
        # =====================================================

        empresa.nombre = (
            valor_o_none(
                fila["Nombre"]
            )
        )

        # =====================================================
        # ARCHIVO INICIAL
        # =====================================================

        empresa.archivo_inicial = (
            valor_o_none(
                fila["Archivo inicial"]
            )
        )

        # =====================================================
        # TELEFONO
        # =====================================================

        empresa.telefono = (
            limpiar_telefono(
                fila["Telefono"]
            )
        )

        # =====================================================
        # EMAIL 1
        # =====================================================

        empresa.email = (
            valor_o_none(
                fila["mail1"]
            )
        )

        # =====================================================
        # EMAIL 2
        # =====================================================

        empresa.email_2 = (
            valor_o_none(
                fila["mail2"]
            )
        )

        # =====================================================
        # EMAIL 3
        # =====================================================

        empresa.email_3 = (
            valor_o_none(
                fila["mail3"]
            )
        )

        # =====================================================
        # PROVINCIA
        # =====================================================

        empresa.provincia = (
            valor_o_none(
                fila["Provincia"]
            )
        )

        # =====================================================
        # COMENTARIOS
        # =====================================================

        empresa.comentarios = (
            valor_o_none(
                fila["Comentarios"]
            )
        )

        # =====================================================
        # GUARDAR
        # =====================================================

        empresa.save(
            update_fields=[
                "nombre",
                "archivo_inicial",
                "telefono",
                "email",
                "email_2",
                "email_3",
                "provincia",
                "comentarios",
            ]
        )

        actualizadas += 1

    return (
        actualizadas,
        no_encontradas,
    )