from decimal import Decimal, InvalidOperation

from django import template


register = template.Library()


@register.filter
def numero_ar(valor, decimales=2):
    """
    Formato argentino:
    1234567.89 -> 1.234.567,89
    """

    if valor is None or valor == "":
        return "-"

    try:
        valor = Decimal(str(valor))

    except (InvalidOperation, ValueError, TypeError):
        return valor

    formato = f"{{:,.{int(decimales)}f}}".format(valor)

    formato = (
        formato
        .replace(",", "@")
        .replace(".", ",")
        .replace("@", ".")
    )

    return formato