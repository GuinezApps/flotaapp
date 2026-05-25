from django import forms
from django.contrib.auth.models import User, Group
from django.contrib.auth.forms import UserCreationForm
from django.core.exceptions import ValidationError

import re

from .models import Vehiculo, CentroCosto


# ============================================================
# Formulario de vehículos
# ============================================================

class VehiculoForm(forms.ModelForm):
    """
    Formulario principal para crear y editar vehículos.

    Incluye una validación importante para la nueva lógica de estados:
    - Si el vehículo está No operativo, debe tener un subestado.
    - Si el vehículo está Operativo o No informado, el subestado se limpia.
    """

    class Meta:

        model = Vehiculo

        exclude = [
            'kilometraje_restante'
        ]

        widgets = {

            'fecha_inicio_contrato':
            forms.DateInput(
                attrs={'type': 'date'}
            ),

            'fecha_termino_contrato':
            forms.DateInput(
                attrs={'type': 'date'}
            ),

            'fecha_baja':
            forms.DateInput(
                attrs={'type': 'date'}
            ),

            'observaciones':
            forms.Textarea(
                attrs={
                    'rows': 4
                }
            ),

        }

    def clean(self):
        """
        Valida coherencia entre estado operacional y subestado.

        Reglas:
        - No operativo requiere subestado.
        - Operativo y No informado no deben conservar subestado.
        """

        cleaned_data = super().clean()

        estado_operacional = cleaned_data.get(
            'estado_operacional'
        )

        subestado_no_operativo = cleaned_data.get(
            'subestado_no_operativo'
        )

        if (
            estado_operacional == 'No operativo'
            and not subestado_no_operativo
        ):

            self.add_error(
                'subestado_no_operativo',
                'Debe indicar el motivo cuando el vehículo está No operativo.'
            )

        if estado_operacional != 'No operativo':

            cleaned_data[
                'subestado_no_operativo'
            ] = None

        return cleaned_data


# ============================================================
# Formulario de creación de usuarios
# ============================================================

class CrearUsuarioForm(UserCreationForm):
    """
    Formulario usado por el rol Master para crear usuarios.

    Crea el usuario base de Django y captura datos adicionales
    que luego se guardan en PerfilUsuario desde la vista.
    """

    rol = forms.ModelChoiceField(
        queryset=Group.objects.filter(
            name__in=[
                'Visualizador',
                'Editor',
                'Master'
            ]
        ),
        required=True,
        label='Permisos'
    )

    email = forms.EmailField(
        required=True,
        label='Correo'
    )

    rut = forms.CharField(
        required=True,
        label='RUT',
        help_text='Formato: 12345678-9'
    )

    nombre = forms.CharField(
        required=True,
        label='Nombre completo'
    )

    telefono = forms.CharField(
        required=True,
        label='Teléfono',
        help_text='Ingrese solo el número. Ej: 974539081'
    )

    cargo = forms.CharField(
        required=True,
        label='Cargo'
    )

    centro_costo = forms.ModelChoiceField(
        queryset=CentroCosto.objects.order_by(
            'codigo'
        ),
        required=True,
        label='Centro de costo'
    )

    sucursal = forms.CharField(
        required=True,
        label='Sucursal'
    )

    class Meta:

        model = User

        fields = [
            'username',
            'email',
            'password1',
            'password2',
            'rut',
            'nombre',
            'telefono',
            'cargo',
            'centro_costo',
            'sucursal',
            'rol'
        ]

    def clean_username(self):
        """
        Evita crear usuarios con username duplicado.
        """

        username = self.cleaned_data.get(
            'username'
        )

        if not username:
            return username

        if User.objects.filter(
            username=username
        ).exists():

            raise ValidationError(
                'Ya existe un usuario con este nombre de usuario.'
            )

        return username

    def clean_email(self):
        """
        Evita crear usuarios con correo duplicado.
        """

        email = self.cleaned_data.get(
            'email'
        )

        if not email:
            return email

        if User.objects.filter(
            email=email
        ).exists():

            raise ValidationError(
                'Ya existe un usuario con este correo.'
            )

        return email

    def clean_rut(self):
        """
        Valida formato de RUT chileno.

        Formato aceptado:
        12345678-9
        12345678-K
        """

        rut = self.cleaned_data.get(
            'rut'
        )

        if not rut:
            return rut

        rut = rut.strip().upper()

        patron = r'^\d{7,8}-[\dK]$'

        if not re.match(
            patron,
            rut
        ):

            raise ValidationError(
                'Formato: 12345678-9'
            )

        return rut

    def clean_telefono(self):
        """
        Normaliza teléfonos chilenos móviles.

        Acepta:
        - 974539081
        - 56974539081
        - +56974539081
        - +56 9 7453 9081

        Guarda:
        +56974539081
        """

        telefono = self.cleaned_data.get(
            'telefono'
        )

        if not telefono:
            return telefono

        telefono = telefono.strip()

        telefono = telefono.replace(
            ' ',
            ''
        ).replace(
            '-',
            ''
        )

        if telefono.startswith(
            '9'
        ):

            telefono = '+56' + telefono

        elif telefono.startswith(
            '569'
        ):

            telefono = '+' + telefono

        patron = r'^\+569\d{8}$'

        if not re.match(
            patron,
            telefono
        ):

            raise ValidationError(
                'Debe ingresar un número chileno válido. Ej: 974539081'
            )

        return telefono
