from django.db import models
from django.contrib.auth.models import User


# ============================================================
# Centros de costo
# ============================================================

class CentroCosto(models.Model):
    codigo = models.CharField(
        max_length=100,
        unique=True
    )

    nombre = models.CharField(
        max_length=150,
        blank=True,
        null=True
    )

    def __str__(self):
        if self.nombre and self.nombre != self.codigo:
            return f"{self.codigo} - {self.nombre}"

        return self.codigo


# ============================================================
# Vehículos
# ============================================================

class Vehiculo(models.Model):
    ESTADOS_ADMINISTRATIVOS = [
        ('Vigente', 'Vigente'),
        ('Dado de Baja', 'Dado de Baja'),
    ]
    ESTADOS_OPERACIONALES = [
        ('Operativo', 'Operativo'),
        ('No operativo', 'No operativo'),
        ('No informado', 'No informado'),
    ]
    SUBESTADOS_NO_OPERATIVO = [
        ('Mantencion', 'En mantención'),
        ('Reparacion', 'En reparación'),
        ('Detenido', 'Detenido'),
    ]

    # --------------------------------------------------------
    # Identificación técnica
    # --------------------------------------------------------

    patente = models.CharField(
        max_length=20,
        unique=True
    )

    anio = models.IntegerField(
        blank=True,
        null=True
    )

    combustible = models.CharField(
        max_length=100,
        blank=True,
        null=True
    )

    traccion = models.CharField(
        max_length=100,
        blank=True,
        null=True
    )

    numero_motor = models.CharField(
        max_length=150,
        blank=True,
        null=True
    )

    numero_chasis_vin = models.CharField(
        max_length=150,
        blank=True,
        null=True
    )

    # --------------------------------------------------------
    # Características del vehículo
    # --------------------------------------------------------

    tipo_vehiculo = models.CharField(
        max_length=100,
        blank=True,
        null=True
    )

    marca = models.CharField(
        max_length=100,
        blank=True,
        null=True
    )

    modelo = models.CharField(
        max_length=150,
        blank=True,
        null=True
    )

    color = models.CharField(
        max_length=100,
        blank=True,
        null=True
    )

    numero_asientos = models.IntegerField(
        blank=True,
        null=True
    )

    # --------------------------------------------------------
    # Datos contractuales / comerciales
    # --------------------------------------------------------

    tarifa_mensual = models.CharField(
        max_length=100,
        blank=True,
        null=True
    )

    tipo_estandar = models.CharField(
        max_length=150,
        blank=True,
        null=True
    )

    fecha_inicio_contrato = models.DateField(
        blank=True,
        null=True
    )

    fecha_termino_contrato = models.DateField(
        blank=True,
        null=True
    )

    periodo_meses = models.IntegerField(
        blank=True,
        null=True
    )

    # --------------------------------------------------------
    # Asignación organizacional
    # --------------------------------------------------------

    centro_costo = models.ForeignKey(
        CentroCosto,
        on_delete=models.PROTECT,
        blank=True,
        null=True
    )

    sucursal = models.CharField(
        max_length=150,
        blank=True,
        null=True
    )

    ciudad = models.CharField(
        max_length=150,
        blank=True,
        null=True
    )

    faena = models.CharField(
        max_length=150,
        blank=True,
        null=True
    )

    gerencia = models.CharField(
        max_length=150,
        blank=True,
        null=True
    )

    # --------------------------------------------------------
    # Kilometraje y mantención
    # --------------------------------------------------------

    limite_kilometraje = models.IntegerField(
        blank=True,
        null=True
    )

    kilometraje_ultima_mantencion = models.IntegerField(
        blank=True,
        null=True
    )

    kilometraje_proxima_mantencion = models.IntegerField(
        blank=True,
        null=True
    )

    kilometraje_actual = models.IntegerField(
        blank=True,
        null=True
    )

    kilometraje_restante = models.IntegerField(
        blank=True,
        null=True
    )

    dias_restantes = models.IntegerField(
        blank=True,
        null=True
    )

    alerta_mantencion = models.CharField(
        max_length=150,
        blank=True,
        null=True
    )

    # --------------------------------------------------------
    # Arrendadora / proveedor
    # --------------------------------------------------------

    empresa_arrendadora = models.CharField(
        max_length=150,
        blank=True,
        null=True
    )

    razon_social = models.CharField(
        max_length=200,
        blank=True,
        null=True
    )

    rut_empresa = models.CharField(
        max_length=50,
        blank=True,
        null=True
    )

    # --------------------------------------------------------
    # Estado administrativo
    # --------------------------------------------------------

    fecha_baja = models.DateField(
        blank=True,
        null=True
    )

    estado_administrativo = models.CharField(
        max_length=100,
        choices=ESTADOS_ADMINISTRATIVOS,
        default='Vigente'
    )

    # --------------------------------------------------------
    # Estado operacional
    # --------------------------------------------------------
    # estado_operacional:
    #   - Operativo
    #   - No operativo
    #   - No informado
    #
    # subestado_no_operativo:
    #   Solo aplica cuando estado_operacional = No operativo.
    # --------------------------------------------------------

    estado_operacional = models.CharField(
        max_length=100,
        choices=ESTADOS_OPERACIONALES,
        default='No informado'
    )

    subestado_no_operativo = models.CharField(
        max_length=50,
        choices=SUBESTADOS_NO_OPERATIVO,
        blank=True,
        null=True
    )

    # --------------------------------------------------------
    # Estado de mantención informado por fuente externa
    # --------------------------------------------------------

    estado_mantencion = models.CharField(
        max_length=100,
        blank=True,
        null=True
    )

    observaciones = models.TextField(
        blank=True,
        null=True
    )

    def __str__(self):
        """
        Representación legible del vehículo en Django Admin,
        formularios y relaciones.
        """

        if self.marca:
            return f"{self.patente} - {self.marca}"

        return self.patente


# ============================================================
# Perfil de usuario
# ============================================================

class PerfilUsuario(models.Model):
    """
    Extiende el usuario nativo de Django con datos internos
    de la empresa.

    User maneja:
    - username
    - email
    - password
    - grupos/permisos

    PerfilUsuario maneja:
    - rut
    - nombre
    - teléfono
    - cargo
    - centro de costo
    - sucursal
    """

    usuario = models.OneToOneField(
        User,
        on_delete=models.CASCADE,
        related_name='perfil'
    )

    rut = models.CharField(
        max_length=20,
        blank=True,
        null=True
    )

    nombre = models.CharField(
        max_length=200,
        blank=True,
        null=True
    )

    telefono = models.CharField(
        max_length=30,
        blank=True,
        null=True
    )

    cargo = models.CharField(
        max_length=150,
        blank=True,
        null=True
    )

    centro_costo = models.ForeignKey(
        CentroCosto,
        on_delete=models.SET_NULL,
        blank=True,
        null=True
    )

    sucursal = models.CharField(
        max_length=150,
        blank=True,
        null=True
    )

    def __str__(self):
        """
        Muestra el nombre del perfil si existe.
        Si no, usa el username del usuario asociado.
        """

        if self.nombre:
            return self.nombre

        return self.usuario.username
        
# ============================================================
# Auditoría / Bitácora
# ============================================================

class BitacoraAccion(models.Model):
    """
    Registra acciones relevantes realizadas dentro del sistema.

    Su objetivo es dejar trazabilidad de:
    - quién realizó una acción
    - cuándo la realizó
    - sobre qué módulo u objeto
    - qué cambio se ejecutó
    """

    ACCIONES = [
        ('CREAR', 'Crear'),
        ('EDITAR', 'Editar'),
        ('ELIMINAR', 'Eliminar'),
        ('CAMBIO_ESTADO', 'Cambio de estado'),
        ('BAJA', 'Dar de baja'),
        ('REACTIVAR', 'Reactivar'),
        ('IMPORTAR', 'Importar datos'),
        ('RESET_PASSWORD', 'Resetear contraseña'),
        ('LOGIN', 'Inicio de sesión'),
        ('LOGOUT', 'Cierre de sesión'),
        ('OTRO', 'Otro'),
    ]

    usuario = models.ForeignKey(
        User,
        on_delete=models.SET_NULL,
        blank=True,
        null=True
    )

    usuario_texto = models.CharField(
        max_length=150,
        blank=True,
        null=True,
        help_text='Nombre de usuario al momento de la acción.'
    )

    accion = models.CharField(
        max_length=50,
        choices=ACCIONES
    )

    modulo = models.CharField(
        max_length=100,
        help_text='Módulo donde ocurrió la acción. Ej: Vehículos, Usuarios, Importación.'
    )

    modelo_afectado = models.CharField(
        max_length=100,
        blank=True,
        null=True,
        help_text='Modelo afectado. Ej: Vehiculo, User, CentroCosto.'
    )

    objeto_id = models.CharField(
        max_length=100,
        blank=True,
        null=True,
        help_text='ID del objeto afectado.'
    )

    objeto_repr = models.CharField(
        max_length=250,
        blank=True,
        null=True,
        help_text='Representación legible del objeto. Ej: QAQA01 - Toyota.'
    )

    descripcion = models.TextField(
        help_text='Descripción resumida de la acción realizada.'
    )

    valor_anterior = models.TextField(
        blank=True,
        null=True,
        help_text='Valor anterior o resumen previo del dato modificado.'
    )

    valor_nuevo = models.TextField(
        blank=True,
        null=True,
        help_text='Valor nuevo o resumen posterior del dato modificado.'
    )

    fecha_hora = models.DateTimeField(
        auto_now_add=True
    )

    ip_origen = models.GenericIPAddressField(
        blank=True,
        null=True
    )

    def __str__(self):

        usuario = self.usuario_texto or 'Sistema'

        return (
            f'{self.fecha_hora:%Y-%m-%d %H:%M} | '
            f'{usuario} | '
            f'{self.get_accion_display()} | '
            f'{self.modulo}'
        )

    class Meta:

        ordering = [
            '-fecha_hora'
        ]

        verbose_name = 'Bitácora de acción'
        verbose_name_plural = 'Bitácora de acciones'