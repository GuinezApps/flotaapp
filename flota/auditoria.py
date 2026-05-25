from .models import BitacoraAccion


def obtener_ip_request(request):
    """
    Obtiene la IP del usuario desde el request.

    Si existe un proxy intermedio, intenta tomar la primera IP
    desde HTTP_X_FORWARDED_FOR.
    """

    x_forwarded_for = request.META.get(
        'HTTP_X_FORWARDED_FOR'
    )

    if x_forwarded_for:

        ip = x_forwarded_for.split(
            ','
        )[0]

    else:

        ip = request.META.get(
            'REMOTE_ADDR'
        )

    return ip


def registrar_bitacora(
    request,
    accion,
    modulo,
    descripcion,
    modelo_afectado=None,
    objeto_id=None,
    objeto_repr=None,
    valor_anterior=None,
    valor_nuevo=None
):
    """
    Registra una acción relevante en la bitácora del sistema.

    Esta función centraliza la creación de registros de auditoría
    para evitar repetir lógica en cada vista.
    """

    usuario = None
    usuario_texto = 'Sistema'

    if request and request.user.is_authenticated:

        usuario = request.user
        usuario_texto = request.user.username

    BitacoraAccion.objects.create(

        usuario=usuario,

        usuario_texto=usuario_texto,

        accion=accion,

        modulo=modulo,

        modelo_afectado=modelo_afectado,

        objeto_id=str(
            objeto_id
        ) if objeto_id else None,

        objeto_repr=objeto_repr,

        descripcion=descripcion,

        valor_anterior=valor_anterior,

        valor_nuevo=valor_nuevo,

        ip_origen=obtener_ip_request(
            request
        ) if request else None

    )