from django.urls import path
from . import views

urlpatterns = [
    path('',views.inicio,name='inicio'),
    path('vehiculos/',views.dashboard,name='dashboard'),
    path('exportar-excel/', views.exportar_vehiculos_excel, name='exportar_vehiculos_excel'),
    path('importar/',views.importar_datos,name='importar_datos'),
    path('confirmar-importacion/',views.confirmar_importacion,name='confirmar_importacion'),
    path('estados/',views.gestionar_estados,name='gestionar_estados'),
    path('vehiculo/<int:id>/',views.detalle_vehiculo,name='detalle_vehiculo'),
    path('vehiculo/nuevo/',views.crear_vehiculo,name='crear_vehiculo'),
    path('vehiculo/<int:id>/editar/',views.editar_vehiculo,name='editar_vehiculo'),
    path('vehiculo/<int:id>/baja/',views.dar_baja_vehiculo,name='dar_baja_vehiculo'),
    path('vehiculo/<int:id>/reactivar/',views.reactivar_vehiculo,name='reactivar_vehiculo'),
    path('usuarios/nuevo/',views.crear_usuario,name='crear_usuario'),
    path('configuracion/',views.configuracion,name='configuracion'),
    path('configuracion/usuarios/',views.usuarios,name='usuarios'),
    path('configuracion/usuarios/<int:id>/editar/',views.editar_usuario,name='editar_usuario'),
    path('configuracion/usuarios/<int:id>/eliminar/',views.eliminar_usuario,name='eliminar_usuario'),
    path('configuracion/usuarios/<int:id>/resetear-password/',views.resetear_password_usuario,name='resetear_password_usuario'),
]