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


# ============================================================
# EMPRESA
# ============================================================

def obtener_empresa(cuit, nombre):

    cuit = limpiar_cuit(cuit)

    if not cuit:
        return None, False

    nombre = valor_o_none(nombre)

    if not nombre:
        nombre = "Empresa sin nombre"

    empresa, creada = Empresa.objects.get_or_create(
        cuit=cuit,
        defaults={
            "nombre": nombre
        }
    )

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

            else:

                valor = valor_o_none(valor)

            datos[campo_modelo] = valor

        # ----------------------------------------------------
        # EMPRESA OFERENTE
        # ----------------------------------------------------

        empresa_oferente, creada = obtener_empresa(
            datos.get("cuit_oferente"),
            datos.get("oferente"),
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