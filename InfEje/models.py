from django.db import models


class Lote(models.Model):
    nombre_archivo = models.CharField(max_length=255, unique=True)
    fecha_carga = models.DateTimeField(auto_now_add=True)
    rubro = models.CharField(max_length=100, default="Servicios de limpieza")

    def __str__(self):
        return self.nombre_archivo


class Empresa(models.Model):
    nombre = models.CharField(max_length=200)
    cuit = models.CharField(max_length=20, unique=True)

    email = models.EmailField(
        blank=True,
        null=True
    )

    email_2 = models.EmailField(
        blank=True,
        null=True
    )

    email_3 = models.EmailField(
        blank=True,
        null=True
    )

    def __str__(self):
        return self.nombre

class RegistroLicitacion(models.Model):
    lote = models.ForeignKey(
        Lote,
        on_delete=models.CASCADE,
        related_name="registros"
    )
    # Empresa asociada
    empresa_oferente = models.ForeignKey(
        Empresa,
        on_delete=models.SET_NULL,
        blank=True,
        null=True,
        related_name="registros_como_oferente"
    )

    empresa_proveedor = models.ForeignKey(
        Empresa,
        on_delete=models.SET_NULL,
        blank=True,
        null=True,
        related_name="registros_como_proveedor"
    )
    # Identificación del proceso
    unidad_negocios = models.CharField(
        max_length=200,
        blank=True,
        null=True
    )

    segmento = models.CharField(
        max_length=200,
        blank=True,
        null=True
    )

    numero_proceso = models.CharField(
        max_length=100,
        blank=True,
        null=True
    )

    portal_compra = models.CharField(
        max_length=100,
        blank=True,
        null=True
    )

    comprador = models.CharField(
        max_length=300,
        blank=True,
        null=True
    )

    estado = models.CharField(
        max_length=100,
        blank=True,
        null=True
    )

    # Renglón
    numero_renglon = models.CharField(
        max_length=50,
        blank=True,
        null=True
    )

    descripcion_renglon = models.TextField(
        blank=True,
        null=True
    )

    cantidad = models.DecimalField(
        max_digits=18,
        decimal_places=4,
        blank=True,
        null=True
    )

    fecha_apertura = models.DateField(
        blank=True,
        null=True
    )

    # Oferente
    oferente = models.CharField(
        max_length=300,
        blank=True,
        null=True
    )

    cuit_oferente = models.CharField(
        max_length=20,
        blank=True,
        null=True
    )

    alternativa = models.CharField(
        max_length=100,
        blank=True,
        null=True
    )

    especificacion_tecnica = models.TextField(
        blank=True,
        null=True
    )

    cantidad_ofertada = models.DecimalField(
        max_digits=18,
        decimal_places=4,
        blank=True,
        null=True
    )

    precio_unitario_oferta = models.DecimalField(
        max_digits=20,
        decimal_places=4,
        blank=True,
        null=True
    )

    precio_total_oferta = models.DecimalField(
        max_digits=20,
        decimal_places=4,
        blank=True,
        null=True
    )

    moneda_oferta = models.CharField(
        max_length=20,
        blank=True,
        null=True
    )

    # Contratación
    numero_oc = models.CharField(
        max_length=100,
        blank=True,
        null=True
    )

    fecha_inicio_contrato = models.DateField(
        blank=True,
        null=True
    )

    fecha_fin_contrato = models.DateField(
        blank=True,
        null=True
    )

    duracion_contrato = models.IntegerField(
        blank=True,
        null=True
    )

    tipo_contrato = models.CharField(
        max_length=100,
        blank=True,
        null=True
    )

    # Proveedor
    proveedor = models.CharField(
        max_length=300,
        blank=True,
        null=True
    )

    cuit_proveedor = models.CharField(
        max_length=20,
        blank=True,
        null=True
    )

    descripcion_contrato = models.TextField(
        blank=True,
        null=True
    )

    cantidad_comprada = models.DecimalField(
        max_digits=18,
        decimal_places=4,
        blank=True,
        null=True
    )

    precio_unitario_compra = models.DecimalField(
        max_digits=20,
        decimal_places=4,
        blank=True,
        null=True
    )

    precio_total_compra = models.DecimalField(
        max_digits=20,
        decimal_places=4,
        blank=True,
        null=True
    )

    moneda_compra = models.CharField(
        max_length=20,
        blank=True,
        null=True
    )

    # Fuente
    url = models.URLField(
        max_length=1000,
        blank=True,
        null=True
    )

    def __str__(self):
        return (
            f"{self.numero_proceso or 'Sin proceso'} - "
            f"{self.oferente or self.proveedor or 'Sin empresa'}"
        )