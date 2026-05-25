from django.db import models
from django.contrib.auth.models import User

class CentroCosto(models.Model):
    codigo = models.CharField(max_length=100, unique=True)
    nombre = models.CharField(max_length=150, blank=True, null=True)

    def __str__(self):

        if self.nombre and self.nombre != self.codigo:

            return f"{self.codigo} - {self.nombre}"

        return self.codigo


class Vehiculo(models.Model):

    ESTADOS_ADMINISTRATIVOS = [
    ('Vigente', 'Vigente'),
    ('Dado de Baja', 'Dado de Baja'),
]

    ESTADOS_OPERACIONALES = [
    ('No informado', 'No informado'),
    ('Operativo', 'Operativo'),
    ('En Mantencion', 'En Mantención'),
    ('FueraServicio', 'Fuera de Servicio'),
]

    patente = models.CharField(max_length=20, unique=True)

    anio = models.IntegerField(blank=True, null=True)
    combustible = models.CharField(max_length=100, blank=True, null=True)
    traccion = models.CharField(max_length=100, blank=True, null=True)
    numero_motor = models.CharField(max_length=150, blank=True, null=True)
    numero_chasis_vin = models.CharField(max_length=150, blank=True, null=True)

    tipo_vehiculo = models.CharField(max_length=100, blank=True, null=True)
    marca = models.CharField(max_length=100, blank=True, null=True)
    modelo = models.CharField(max_length=150, blank=True, null=True)
    color = models.CharField(max_length=100, blank=True, null=True)
    numero_asientos = models.IntegerField(blank=True, null=True)

    tarifa_mensual = models.CharField(max_length=100,blank=True, null=True)
    centro_costo = models.ForeignKey(
        CentroCosto,
        on_delete=models.PROTECT,
        blank=True,
        null=True
    )

    tipo_estandar = models.CharField(max_length=150, blank=True, null=True)
    fecha_inicio_contrato = models.DateField(blank=True, null=True)
    fecha_termino_contrato = models.DateField(blank=True, null=True)
    periodo_meses = models.IntegerField(blank=True, null=True)

    limite_kilometraje = models.IntegerField(blank=True, null=True)
    kilometraje_ultima_mantencion = models.IntegerField(blank=True, null=True)
    kilometraje_proxima_mantencion = models.IntegerField(blank=True, null=True)
    kilometraje_actual = models.IntegerField(blank=True, null=True)
    kilometraje_restante = models.IntegerField(blank=True, null=True)
    dias_restantes = models.IntegerField(blank=True, null=True)

    alerta_mantencion = models.CharField(max_length=150, blank=True, null=True)
    empresa_arrendadora = models.CharField(max_length=150, blank=True, null=True)
    razon_social = models.CharField(max_length=200, blank=True, null=True)
    rut_empresa = models.CharField(max_length=50, blank=True, null=True)

    sucursal = models.CharField(max_length=150, blank=True, null=True)
    ciudad = models.CharField(max_length=150, blank=True, null=True)
    faena = models.CharField(max_length=150, blank=True, null=True)
    gerencia = models.CharField(max_length=150, blank=True, null=True)

    fecha_baja = models.DateField(blank=True, null=True)
    
    estado_administrativo = models.CharField(
        max_length=100,
        choices=ESTADOS_ADMINISTRATIVOS,
        default='Vigente'
    )

    estado_operacional = models.CharField(
        max_length=100,
        choices=ESTADOS_OPERACIONALES,
        default='No informado'
    )

    estado_mantencion = models.CharField(
        max_length=100,
        blank=True,
        null=True
    )
    observaciones = models.TextField(blank=True, null=True)

    def __str__(self):
        if self.marca:
            return f"{self.patente} - {self.marca}"
        return self.patente

class PerfilUsuario(models.Model):

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

        if self.nombre:
            return self.nombre

        return self.usuario.username