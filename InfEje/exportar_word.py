from io import BytesIO
from PIL import Image, ImageDraw, ImageFont

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from decimal import Decimal, ROUND_HALF_UP

def obtener_fuente(tamano, negrita=False):
    """
    Busca una fuente disponible en Windows o Linux.
    """

    rutas = []

    if negrita:
        rutas.extend([
            r"C:\Windows\Fonts\arialbd.ttf",
            "/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
        ])
    else:
        rutas.extend([
            r"C:\Windows\Fonts\arial.ttf",
            "/usr/share/fonts/truetype/msttcorefonts/Arial.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
        ])

    for ruta in rutas:
        try:
            return ImageFont.truetype(ruta, tamano)
        except Exception:
            pass

    return ImageFont.load_default()


def texto(valor):
    """
    Convierte cualquier valor a texto limpio.
    """
    if valor is None:
        return "-"

    valor = str(valor).strip()

    return valor if valor else "-"


def formatear_fecha(fecha):
    if not fecha:
        return "-"

    return fecha.strftime("%d/%m/%Y")


def obtener_empresa_registro(registro):
    """
    Devuelve la Empresa asociada al registro.
    """

    if registro.empresa_oferente:
        return registro.empresa_oferente

    if registro.empresa_proveedor:
        return registro.empresa_proveedor

    return None

def agregar_tabla_resultados(documento, registros, empresa):
    """
    Genera la tabla de resultados como imagen
    y la inserta en el documento Word.

    Los anchos son fijos.
    Los competidores se muestran TODOS,
    separados por comas y sin recortar.
    """

    columnas = [
        ("Proceso", 200),
        ("R", 35),
        ("OC", 45),
        ("Comprador", 380),
        ("Empresas que también licitaron", 700),
        ("Apertura", 110),
        ("Estado", 120),
        ("Oferta", 120),
        ("Mon", 50),
        ("Descripción", 320),
    ]
    

    columnas = [
        (nombre, ancho )
        for nombre, ancho in columnas
    ]
    # ========================================================
    # FUENTES
    # ========================================================

    fuente_encabezado = obtener_fuente(
        20 ,
        negrita=True
    )

    fuente_normal = obtener_fuente(
        19 
    )

    # SOLO esta columna lleva letra más chica
    fuente_competidores = obtener_fuente(
        18 
    )

    # ========================================================
    # PREPARAR FILAS
    # ========================================================

    filas = []

    for registro in registros:

        oc = (
            registro.numero_oc
            if registro.numero_oc
            else "-"
        )

        oferta = (
            f"{registro.precio_total_oferta:,.2f}"
            if registro.precio_total_oferta is not None
            else "-"
        )

        # ----------------------------------------------------
        # COMPETIDORES
        # ----------------------------------------------------

        competidores = getattr(
            registro,
            "competidores",
            None
        )

        if competidores:

            if isinstance(
                competidores,
                (list, tuple, set)
            ):
                competidores = ", ".join(
                    str(nombre).strip()
                    for nombre in competidores
                    if str(nombre).strip()
                )

            else:
                competidores = str(
                    competidores
                ).strip()

                competidores = (
                    competidores
                    .replace(" ; ", ", ")
                    .replace(";", ", ")
                )

        else:
            competidores = "-"

        # ----------------------------------------------------
        # FILA
        # ----------------------------------------------------

        filas.append([
            texto(registro.numero_proceso),
            texto(registro.numero_renglon),
            texto(oc),
            texto(registro.comprador),
            competidores,
            formatear_fecha(
                registro.fecha_apertura
            ),
            texto(registro.estado),
            oferta,
            texto(registro.moneda_oferta),
            texto(registro.descripcion_renglon),
        ])

    # ========================================================
    # CONFIGURACIÓN
    # ========================================================

    margen = 14  
    alto_linea = 22 
    alto_encabezado = 30 

    ancho_total = (
        sum(
            ancho
            for _, ancho in columnas
        )
        + margen * 1
    )

    # ========================================================
    # IMAGEN PROVISORIA PARA MEDIR TEXTO
    # ========================================================

    imagen_provisoria = Image.new(
        "RGB",
        (
            ancho_total,
            100
        ),
        "white"
    )

    dibujo = ImageDraw.Draw(
        imagen_provisoria
    )

    # ========================================================
    # ENVOLVER TEXTO
    # ========================================================

    def envolver_texto(
        valor,
        fuente,
        ancho_disponible
    ):
        valor = str(valor)

        if not valor:
            return ["-"]

        palabras = valor.split()
        lineas = []
        linea_actual = ""

        for palabra in palabras:

            if not linea_actual:
                candidata = palabra
            else:
                candidata = (
                    linea_actual
                    + " "
                    + palabra
                )

            bbox = dibujo.textbbox(
                (0, 0),
                candidata,
                font=fuente
            )

            ancho_texto = (
                bbox[2] - bbox[0]
            )

            if ancho_texto <= ancho_disponible:
                linea_actual = candidata
            else:

                if linea_actual:
                    lineas.append(
                        linea_actual
                    )

                linea_actual = palabra

        if linea_actual:
            lineas.append(
                linea_actual
            )

        return lineas or ["-"]


    
    def envolver_competidores(
        valor,
        ancho_disponible
    ):
        """
        Muestra TODOS los competidores,
        separados por comas y acomodados
        en un máximo de 2 líneas.

        Reduce solamente el tamaño de letra
        si es necesario para que entren.
        """

        valor = str(valor).strip()

        if not valor:
            return [("-", fuente_competidores)]

        # --------------------------------------------------------
        # SEPARAR EMPRESAS
        # --------------------------------------------------------

        empresas = [
            empresa.strip()
            for empresa in valor.split(",")
            if empresa.strip()
        ]

        if not empresas:
            return [("-", fuente_competidores)]

        # --------------------------------------------------------
        # USAR SIEMPRE LA MISMA FUENTE
        # --------------------------------------------------------

        

        fuente = fuente_competidores


        linea_1 = ""
        linea_2 = ""

        for empresa in empresas:

            candidata_1 = (
                empresa
                if not linea_1
                else linea_1 + ", " + empresa
            )

            ancho_1 = dibujo.textbbox(
                (0, 0),
                candidata_1,
                font=fuente
            )[2]

            if ancho_1 <= ancho_disponible:
                linea_1 = candidata_1
                continue

            candidata_2 = (
                empresa
                if not linea_2
                else linea_2 + ", " + empresa
            )

            linea_2 = candidata_2

        lineas = []

        if linea_1:
            lineas.append(
                (linea_1, fuente)
            )

        if linea_2:
            lineas.append(
                (linea_2, fuente)
            )

        return lineas
        
    def texto_una_linea(valor, fuente, ancho_disponible):
        """
        Muestra el texto en una sola línea.
        Si no entra, lo corta y agrega ...
        """

        valor = str(valor).strip()

        if not valor:
            return "-"

        # Sacamos saltos de línea y espacios repetidos
        valor = " ".join(valor.split())

        bbox = dibujo.textbbox(
            (0, 0),
            valor,
            font=fuente
        )

        ancho_texto = bbox[2] - bbox[0]

        if ancho_texto <= ancho_disponible:
            return valor

        sufijo = "..."

        texto_cortado = ""

        for caracter in valor:

            candidata = (
                texto_cortado
                + caracter
                + sufijo
            )

            bbox = dibujo.textbbox(
                (0, 0),
                candidata,
                font=fuente
            )

            ancho_candidata = (
                bbox[2] - bbox[0]
            )

            if ancho_candidata > ancho_disponible:
                break

            texto_cortado += caracter

        return texto_cortado.rstrip() + sufijo

    # ========================================================
    # CALCULAR ALTURA DE CADA FILA
    # ========================================================

    # ========================================================
    # CALCULAR ALTURA DE CADA FILA
    # ========================================================

    alturas_filas = []

    for fila in filas:

        # --------------------------------------------
        # Competidores: SIEMPRE máximo 2 líneas
        # --------------------------------------------

        competidores = fila[4]

        lineas_competidores = envolver_competidores(
            competidores,
            columnas[4][1] - 12
        )

        cantidad_lineas = 1

        # Los competidores pueden ocupar hasta 2 líneas
        cantidad_lineas = max(
            cantidad_lineas,
            len(lineas_competidores)
        )

        # --------------------------------------------
        # Resto de las columnas
        # --------------------------------------------

        for indice, valor in enumerate(fila):

            # Competidores ya los calculamos arriba
            if indice == 4:
                continue

            _, ancho = columnas[indice]

            # Descripción: una sola línea
            if indice == 9:
                cantidad_lineas = max(
                    cantidad_lineas,
                    1
                )
                continue

            lineas = envolver_texto(
                valor,
                fuente_normal,
                ancho - 12
            )

            cantidad_lineas = max(
                cantidad_lineas,
                len(lineas)
            )

        # --------------------------------------------
        # Altura de la fila
        # --------------------------------------------

        alturas_filas.append(
            cantidad_lineas * alto_linea
        )



    # ========================================================
    # ALTURA TOTAL
    # ========================================================

    alto_total = (
         margen
         + alto_encabezado
         + sum(alturas_filas)
         + margen
    )

    # ========================================================
    # CREAR IMAGEN DEFINITIVA
    # ========================================================

    imagen = Image.new(
        "RGB",
        (
            ancho_total,
            alto_total
        ),
        "white"
    )

    dibujo = ImageDraw.Draw(
        imagen
    )

    # ========================================================
    # ENCABEZADOS
    # ========================================================

    x = margen
    y = margen

    for nombre_columna, ancho in columnas:

        dibujo.rectangle(
            [
                (x, y),
                (
                    x + ancho,
                    y + alto_encabezado
                )
            ],
            outline="black",
            fill="#EAEAEA"
        )

        lineas = envolver_texto(
            nombre_columna,
            fuente_encabezado,
            ancho - 10
        )

        y_texto = (
            y
            + (
                alto_encabezado
                - len(lineas) * alto_linea
            ) / 2
        )

        for linea in lineas:

            dibujo.text(
                (
                    x + 5,
                    y_texto
                ),
                linea,
                font=fuente_encabezado,
                fill="black"
            )

            y_texto += alto_linea

        x += ancho

    y += alto_encabezado

    # ========================================================
    # FILAS
    # ========================================================

    for numero_fila, fila in enumerate(filas):

        alto_fila = alturas_filas[
            numero_fila
        ]

        x = margen

        for indice, valor in enumerate(fila):

            _, ancho = columnas[indice]

            # Competidores con letra más chica
            if indice == 4:

                lineas_competidores = envolver_competidores(
                    valor,
                    ancho - 12
                )

                lineas = [
                    linea
                    for linea, fuente
                    in lineas_competidores
                ]

                fuente = (
                    lineas_competidores[0][1]
                    if lineas_competidores
                    else fuente_competidores
                )

            elif indice == 9:

                # ============================================
                # DESCRIPCIÓN: UNA SOLA LÍNEA + ...
                # ============================================

                fuente = fuente_normal

                descripcion = texto_una_linea(
                    valor,
                    fuente,
                    ancho - 12
                )

                lineas = [
                    descripcion
                ]

            else:

                fuente = fuente_normal

                lineas = envolver_texto(
                    valor,
                    fuente,
                    ancho - 12
                )

            dibujo.rectangle(
                [
                    (x, y),
                    (
                        x + ancho,
                        y + alto_fila
                    )
                ],
                outline="black"
            )

            y_texto = (
                y
                + (
                    alto_fila
                    - len(lineas) * alto_linea
                ) / 2
            )

            for linea in lineas:

                dibujo.text(
                    (
                        x + 6,
                        y_texto
                    ),
                    linea,
                    font=fuente,
                    fill="black"
                )

                y_texto += alto_linea

            x += ancho

        y += alto_fila

    # ========================================================
    # RECORTAR LA IMAGEN AL ALTO REAL DE LA TABLA
    # ========================================================

    alto_real = y + margen

    imagen = imagen.crop(
        (
            0,
            0,
            ancho_total,
            alto_real
        )
    )
    # ========================================================
    # GUARDAR IMAGEN EN MEMORIA
    # ========================================================

    buffer = BytesIO()

    imagen.save(
        buffer,
        format="PNG",
        dpi=(300, 300),
        optimize=True
    )

    buffer.seek(0)

    # ========================================================
    # INSERTAR IMAGEN EN WORD
    # ========================================================

    parrafo = documento.add_paragraph()

    parrafo.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    run = parrafo.add_run()

    run.add_picture(
        buffer,
        width=Inches(8.0)
    )
def agregar_linea(documento, etiqueta, valor, tamano=None):
    """
    Agrega una línea:
    Etiqueta: valor
    """

    parrafo = documento.add_paragraph()

    parrafo.paragraph_format.space_after = Pt(0)

    run = parrafo.add_run(
        f"{etiqueta}: "
    )

    run.bold = True

    run_valor = parrafo.add_run(
        texto(valor)
    )

    if tamano is not None:
        run.font.size = Pt(tamano)
        run_valor.font.size = Pt(tamano)

    return parrafo

    
def calcular_cotizacion(registro, porcentaje_sa):
    """
    Calcula la cotización de Mantenimiento de Oferta
    tomando solamente un registro.

    Reglas actuales:
        - SA = 5% de la oferta
        - Tasa = 1%
        - Período = trimestral (4)
        - Prima = fórmula PRIMA del cotizador
        - Premio simple = fórmula PREMIO FINAL del cotizador
    """

    if registro is None:
        return None

    if registro.precio_total_oferta is None:
        return None

    oferta = Decimal(
        str(registro.precio_total_oferta)
    )


    suma_asegurada = (
        oferta * Decimal(str(porcentaje_sa))
        / Decimal("100")
    )
    tasa = Decimal("1")
    periodos = Decimal("4")

    moneda = (
        registro.moneda_oferta
        or ""
    ).strip().upper()

    # ========================================================
    # COTIZADOR USD
    # ========================================================

    if "USD" in moneda or "U$S" in moneda or "$US" in moneda:

        # PRIMA MINIMA USD
        prima_minima = (
            Decimal("17500")
            / Decimal("1510")
        )

        # DERECHO DE EMISION USD
        derecho_emision = (
            Decimal("15500")
            / Decimal("1510")
        )

        prima_calculada = (
            suma_asegurada
            * tasa
            / Decimal("100")
            / periodos
        )

        # Fórmula PRIMA del Excel
        prima = max(
            prima_calculada,
            prima_minima
        )

        # Fórmula PREMIO FINAL del Excel
        recargo_administrativo = (
            prima * Decimal("0.10")
        )

        subtotal = (
            prima
            + derecho_emision
            + recargo_administrativo
        )

        intereses_internos = (
            subtotal * Decimal("0.001")
        )

        tasa_ssn = (
            subtotal * Decimal("0.006")
        )

        osseg = (
            subtotal * Decimal("0.005")
        )

        sellos = (
            subtotal * Decimal("0.0171")
        )

        iva = (
            subtotal * Decimal("0.21")
        )

        total_impuestos = (
            intereses_internos
            + tasa_ssn
            + osseg
            + sellos
            + iva
        )

        premio_final = (
            subtotal
            + total_impuestos
        )

        moneda_texto = "USD"

    # ========================================================
    # COTIZADOR PESOS
    # ========================================================

    else:

        # PRIMA MINIMA $
        prima_minima = Decimal("18000")

        # DERECHO DE EMISION $
        derecho_emision = Decimal("15500")

        prima_calculada = (
            suma_asegurada
            * tasa
            / Decimal("100")
            / periodos
        )

        # Fórmula PRIMA del Excel
        prima = max(
            prima_calculada,
            prima_minima
        )

        # Fórmula PREMIO FINAL del Excel
        escribania = Decimal("18500")
        colegio = Decimal("16500")

        recargo_administrativo = (
            prima * Decimal("0.10")
        )

        subtotal = (
            prima
            + derecho_emision
            + escribania
            + colegio
            + recargo_administrativo
        )

        intereses_internos = (
            subtotal * Decimal("0.001")
        )

        tasa_ssn = (
            subtotal * Decimal("0.006")
        )

        osseg = (
            subtotal * Decimal("0.005")
        )

        sellos = (
            subtotal * Decimal("0.0171")
        )

        iva = (
            subtotal * Decimal("0.21")
        )

        total_impuestos = (
            intereses_internos
            + tasa_ssn
            + osseg
            + sellos
            + iva
        )

        premio_final = (
            subtotal
            + total_impuestos
        )

        moneda_texto = "$"

    # ========================================================
    # REDONDEO
    # ========================================================

    def redondear(valor):
        return valor.quantize(
            Decimal("0.01"),
            rounding=ROUND_HALF_UP
        )

    return {
        "oferta": redondear(oferta),
        "suma_asegurada": redondear(
            suma_asegurada
        ),
        "prima": redondear(prima),
        "premio_final": redondear(
            premio_final
        ),
        "moneda": moneda_texto,
        "proceso": texto(
            registro.numero_proceso
        ),
        "renglon": texto(
            registro.numero_renglon
        ),
    }

def generar_word_mails(empresas_con_registros):
    """
    Genera un único Word.

    empresas_con_registros:
        [
            (empresa, [registro1, registro2, ...]),
            ...
        ]
    """

    documento = Document()

    # --------------------------------------------------------
    # CONFIGURACIÓN DE PÁGINA
    # --------------------------------------------------------

    seccion = documento.sections[0]

    seccion.top_margin = Inches(0.55)
    seccion.bottom_margin = Inches(0.55)
    seccion.left_margin = Inches(0.55)
    seccion.right_margin = Inches(0.55)

    for indice, (
        empresa,
        registros
    ) in enumerate(empresas_con_registros):

        # ----------------------------------------------------
        # NUEVA PÁGINA
        # ----------------------------------------------------

        if indice > 0:
            documento.add_page_break()

        # ----------------------------------------------------
        # DESTINATARIO
        # ----------------------------------------------------

        nombre_pila = (
            empresa.nombre_pila
            or empresa.nombre
            or ""
        ).strip()

        tipo = (
            empresa.tipo_destinatario
            or "empresa"
        ).lower()

        if tipo == "persona":

            destinatario = (
                f"Hola {nombre_pila}, "
                f"¿Cómo estás?"
            )


            texto_mail = (
                f"Hola {nombre_pila}, ¿Cómo estás?\n\n"
                "Siguiendo con el contacto que tuvimos "
                "oportunamente, y a modo de ejemplo, te "
                "acercamos una simulación de cotización de "
                "las Cauciones de Mantenimiento de Oferta\n"
                "correspondientes a algunos procesos "
                "que identificamos a partir de información "
                "pública.\n\n"
                "En este caso, te mostramos cuánto hubiera "
                "sido nuestra cotización para una de esas MO.\n\n"
                "Sujeto, naturalmente, a las condiciones de "
                "emisión y evaluación de la compañía.\n\n"
                "La idea es que tengas una referencia concreta "
                "para próximas licitaciones.\n\n"
                "Por otro lado, si me enviás la póliza que "
                "contrataste oportunamente, podemos compararla "
                "y evaluar la posibilidad de mejorar el costo "
                "hasta un 30%.\n\n"
                "Saludos,\n"
                "Andrés"
            )
        else:

            destinatario = (
                f"Estimados {nombre_pila},"
            )


            texto_mail = (
                f"Estimados {nombre_pila},\n\n"
                "Siguiendo con el contacto que tuvimos "
                "oportunamente, y a modo de ejemplo, les "
                "acercamos una simulación de cotización de "
                "las Cauciones de Mantenimiento de Oferta\n"
                "correspondientes a algunos procesos "
                "que identificamos a partir de información "
                "pública.\n\n"
                "En este caso, les mostramos cuánto hubiera "
                "sido nuestra cotización para una de esas MO.\n"
                "Sujeto, naturalmente, a las condiciones de "
                "emisión y evaluación de la compañía.\n\n\n"
                "La idea es que tengan una referencia concreta "
                "para próximas licitaciones.\n"
                "Por otro lado, si nos envían la póliza que "
                "contrataron oportunamente, podemos compararla "
                "y evaluar la posibilidad de mejorar el costo "
                "hasta un 30%.\n\n"
                "Saludos,\n"
                "Andrés"
            )

        mails_destinatario = [
            mail
            for mail in [
                empresa.email,
                empresa.email_2,
                empresa.email_3,
            ]
            if mail
        ]

        mails_destinatario = ", ".join(mails_destinatario)


        # ----------------------------------------------------
        # DESTINATARIO
        # ----------------------------------------------------

        agregar_linea(
            documento,
            "Destinatario",
            mails_destinatario
        )

        # ----------------------------------------------------
        # ASUNTO
        # ----------------------------------------------------

        agregar_linea(
            documento,
            "Asunto",
            "Simulación de cotización de Cauciones de MO"
        )

        # ----------------------------------------------------
        # TEXTO
        # ----------------------------------------------------

        parrafo = documento.add_paragraph()

        run = parrafo.add_run(
            ""
        )

        run.bold = True

        # ----------------------------------------------------
        # CUERPO DEL MAIL
        # ----------------------------------------------------

        for linea in texto_mail.split("\n"):

            if not linea.strip():
                continue

            parrafo = documento.add_paragraph()

            parrafo.paragraph_format.space_after = Pt(6)

            run = parrafo.add_run(linea)

            run.font.size = Pt(11)

            if linea.startswith("Estimados ") or linea.startswith("Hola ") or linea.startswith("correspondientes ") or linea.startswith("Sujeto") or linea.startswith("Por otro lado") :
                parrafo.paragraph_format.space_after = Pt(12)





        # ----------------------------------------------------
        # RESULTADOS
        # ----------------------------------------------------



        agregar_tabla_resultados(
            documento,
            registros,
            empresa
        )
        # ----------------------------------------------------
        # COTIZACIÓN MANTENIMIENTO DE OFERTA
        # ----------------------------------------------------

        # ----------------------------------------------------
        # COTIZACIÓN
        # ----------------------------------------------------

        registro_cotizacion = next(
            (
                registro
                for registro in registros
                if registro.precio_total_oferta is not None
            ),
            None
        )

        cotizacion = calcular_cotizacion(
            registro_cotizacion,
            5
        )

        if cotizacion:

            # Título
            parrafo = documento.add_paragraph()
            run = parrafo.add_run("Cotización")
            run.bold = True
            run.font.size = Pt(9)

            # Proceso + Renglón + Oferta
            parrafo = documento.add_paragraph()

            run = parrafo.add_run("Proceso: ")
            run.bold = True
            run.font.size = Pt(9)

            run = parrafo.add_run(
                f'{cotizacion["proceso"]}  /  '
            )
            run.font.size = Pt(9)

            run = parrafo.add_run("Renglón: ")
            run.bold = True
            run.font.size = Pt(9)

            run = parrafo.add_run(
                f'{cotizacion["renglon"]}  /  '
            )
            run.font.size = Pt(9)

            run = parrafo.add_run("Oferta: ")
            run.bold = True
            run.font.size = Pt(9)

            run = parrafo.add_run(
                f'{cotizacion["moneda"]} '
                f'{cotizacion["oferta"]:,.2f}'
            )
            run.font.size = Pt(9)

            # Mantenimiento de Oferta
            agregar_linea(
                documento,
                "— Mantenimiento de Oferta 5%",
                "-"
            ).runs[0].font.size = Pt(9)

            agregar_linea(
                documento,
                "Suma Asegurada (5%)",
                f'{cotizacion["moneda"]} '
                f'{cotizacion["suma_asegurada"]:,.2f}'
            ).runs[0].font.size = Pt(9)

            agregar_linea(
                documento,
                "Prima neta",
                f'{cotizacion["moneda"]} '
                f'{cotizacion["prima"]:,.2f}'
            ).runs[0].font.size = Pt(9)

            agregar_linea(
                documento,
                "Premio simple",
                f'{cotizacion["moneda"]} '
                f'{cotizacion["premio_final"]:,.2f}'
            ).runs[0].font.size = Pt(9)

            # Línea en blanco entre cotizaciones
            documento.add_paragraph()

            # Adjudicación 10%
            cotizacion_adjudicacion = calcular_cotizacion(
                registro_cotizacion,
                10
            )

            if cotizacion_adjudicacion:

                agregar_linea(
                    documento,
                    "— Adjudicación 10%",
                    "-"
                ).runs[0].font.size = Pt(9)

                agregar_linea(
                    documento,
                    "Suma Asegurada (10%)",
                    f'{cotizacion_adjudicacion["moneda"]} '
                    f'{cotizacion_adjudicacion["suma_asegurada"]:,.2f}'
                ).runs[0].font.size = Pt(9)

                agregar_linea(
                    documento,
                    "Prima neta",
                    f'{cotizacion_adjudicacion["moneda"]} '
                    f'{cotizacion_adjudicacion["prima"]:,.2f}'
                ).runs[0].font.size = Pt(9)

                agregar_linea(
                    documento,
                    "Premio simple",
                    f'{cotizacion_adjudicacion["moneda"]} '
                    f'{cotizacion_adjudicacion["premio_final"]:,.2f}'
                ).runs[0].font.size = Pt(9)

    # --------------------------------------------------------
    # GUARDAR
    # --------------------------------------------------------

    buffer = BytesIO()

    documento.save(buffer)

    buffer.seek(0)

    return buffer