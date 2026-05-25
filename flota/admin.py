from django.contrib import admin
from .models import CentroCosto, Vehiculo


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