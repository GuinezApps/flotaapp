from django import forms
from django.contrib.auth.models import User, Group
from django.contrib.auth.forms import UserCreationForm
from django.core.exceptions import ValidationError

import re

from .models import Vehiculo, PerfilUsuario, CentroCosto


class VehiculoForm(forms.ModelForm):

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
            )

        }


class CrearUsuarioForm(UserCreationForm):

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

        rut = self.cleaned_data.get(
            'rut'
        )

        if not rut:
            return rut

        rut = rut.strip().upper()

        patron = r'^\d{7,8}-[\dkK]$'

        if not re.match(
            patron,
            rut
        ):

            raise ValidationError(
                'Formato: 12345678-9'
            )

        return rut


    def clean_telefono(self):

        telefono = self.cleaned_data.get(
            'telefono'
        )

        if not telefono:
            return telefono

        # eliminar espacios y guiones
        telefono = telefono.replace(
            ' ',
            ''
        ).replace(
            '-',
            ''
        )

        # si escribe solo el número
        if telefono.startswith('9'):

            telefono = '+56' + telefono

        # si escribe 569...
        elif telefono.startswith('569'):

            telefono = '+' + telefono

        import re

        patron = r'^\+569\d{8}$'

        if not re.match(
            patron,
            telefono
        ):

            raise ValidationError(
                'Debe ingresar un número chileno válido. Ej: 974539081'
            )

        return telefono