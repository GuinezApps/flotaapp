from django.contrib import admin
from .models import Vehiculo, CentroCosto, PerfilUsuario, BitacoraAccion

@admin.register(CentroCosto)
class CentroCostoAdmin(admin.ModelAdmin):
    list_display = ('codigo', 'nombre')
    search_fields = ('codigo', 'nombre')


@admin.register(Vehiculo)
class VehiculoAdmin(admin.ModelAdmin):
    list_display = (
        'patente',
        'marca',
        'modelo',
        'mostrar_anio',
        'estado_administrativo',
        'estado_operacional',
        'estado_mantencion',
        'centro_costo',
        'kilometraje_actual',
    )

    list_filter = (
        'estado_administrativo',
        'estado_operacional',
        'estado_mantencion',
        'centro_costo',
        'marca',
    )

    search_fields = (
        'patente',
        'marca',
        'modelo',
        'numero_chasis_vin',
        'numero_motor',
    )

    @admin.display(description='Año')
    def mostrar_anio(self, obj):
        return obj.anio

@admin.register(BitacoraAccion)
class BitacoraAccionAdmin(admin.ModelAdmin):

    list_display = [
        'fecha_hora',
        'usuario_texto',
        'accion',
        'modulo',
        'modelo_afectado',
        'objeto_repr',
    ]

    list_filter = [
        'accion',
        'modulo',
        'modelo_afectado',
        'fecha_hora',
    ]

    search_fields = [
        'usuario_texto',
        'descripcion',
        'objeto_repr',
        'valor_anterior',
        'valor_nuevo',
    ]

    readonly_fields = [
        'usuario',
        'usuario_texto',
        'accion',
        'modulo',
        'modelo_afectado',
        'objeto_id',
        'objeto_repr',
        'descripcion',
        'valor_anterior',
        'valor_nuevo',
        'fecha_hora',
        'ip_origen',
    ]