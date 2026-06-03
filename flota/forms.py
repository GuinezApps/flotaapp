from django import forms
from django.contrib.auth.models import User, Group
from django.contrib.auth.forms import UserCreationForm
from django.core.exceptions import ValidationError
from .models import Vehiculo, CentroCosto, MantencionVehiculo
import re

# ============================================================
# Formulario de vehículos
# ============================================================

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
            ),

        }
        
        labels = {
            'anio': 'Año',
        }
        
    def clean_patente(self):

        patente = self.cleaned_data.get(
            'patente'
        )

        if not patente:
            return patente

        patente = patente.strip().upper()

        return patente

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
        
class MantencionVehiculoForm(forms.ModelForm):

    class Meta:

        model = MantencionVehiculo

        fields = [
            'tipo_mantencion',
            'estado',
            'fecha_programada',
            'kilometraje_programado',
            'fecha_ingreso',
            'kilometraje_ingreso',
            'motivo',
            'observacion',
        ]

        widgets = {
            'tipo_mantencion': forms.Select(
                attrs={
                    'class': 'form-control'
                }
            ),

            'estado': forms.Select(
                attrs={
                    'class': 'form-control'
                }
            ),

            'fecha_programada': forms.DateInput(
                attrs={
                    'type': 'date',
                    'class': 'form-control'
                }
            ),

            'kilometraje_programado': forms.NumberInput(
                attrs={
                    'class': 'form-control',
                    'placeholder': 'Ejemplo: 90000'
                }
            ),

            'fecha_ingreso': forms.DateInput(
                attrs={
                    'type': 'date',
                    'class': 'form-control'
                }
            ),

            'kilometraje_ingreso': forms.NumberInput(
                attrs={
                    'class': 'form-control',
                    'placeholder': 'Kilometraje al ingresar a mantención'
                }
            ),

            'motivo': forms.TextInput(
                attrs={
                    'class': 'form-control',
                    'placeholder': 'Ejemplo: mantención preventiva, cambio de aceite, siniestro, revisión programada...'
                }
            ),

            'observacion': forms.Textarea(
                attrs={
                    'class': 'form-control',
                    'rows': 4,
                    'placeholder': 'Observación o detalle de la mantención...'
                }
            ),
        }

    def clean(self):

        cleaned_data = super().clean()

        tipo_mantencion = cleaned_data.get(
            'tipo_mantencion'
        )

        fecha_programada = cleaned_data.get(
            'fecha_programada'
        )

        kilometraje_programado = cleaned_data.get(
            'kilometraje_programado'
        )

        if tipo_mantencion == 'PROGRAMADA' and not fecha_programada:

            self.add_error(
                'fecha_programada',
                'Debes indicar una fecha programada para este tipo de mantención.'
            )

        if tipo_mantencion == 'KILOMETRAJE' and not kilometraje_programado:

            self.add_error(
                'kilometraje_programado',
                'Debes indicar el kilometraje programado para este tipo de mantención.'
            )

        return cleaned_data

class CerrarMantencionVehiculoForm(forms.Form):

    ESTADOS_POSTERIORES = [
        ('NO_CAMBIAR', 'No cambiar estado del vehículo'),
        ('OPERATIVO', 'Marcar como Operativo'),
        ('NO_INFORMADO', 'Marcar como No informado'),
        ('NO_OPERATIVO_MANTENCION', 'Mantener como No operativo / Mantención'),
        ('NO_OPERATIVO_REPARACION', 'Mantener como No operativo / Reparación'),
        ('NO_OPERATIVO_DETENIDO', 'Marcar como No operativo / Detenido'),
    ]

    fecha_cierre = forms.DateField(
        widget=forms.DateInput(
            attrs={
                'type': 'date',
                'class': 'form-control'
            }
        ),
        label='Fecha de cierre'
    )

    kilometraje_cierre = forms.IntegerField(
        required=False,
        widget=forms.NumberInput(
            attrs={
                'class': 'form-control',
                'placeholder': 'Kilometraje al cierre de la mantención'
            }
        ),
        label='Kilometraje de cierre'
    )

    estado_posterior_vehiculo = forms.ChoiceField(
        choices=ESTADOS_POSTERIORES,
        initial='NO_CAMBIAR',
        widget=forms.Select(
            attrs={
                'class': 'form-control'
            }
        ),
        label='Estado posterior del vehículo'
    )

    observacion_cierre = forms.CharField(
        required=False,
        widget=forms.Textarea(
            attrs={
                'class': 'form-control',
                'rows': 4,
                'placeholder': 'Observación de cierre de la mantención...'
            }
        ),
        label='Observación de cierre'
    )

    def __init__(self, *args, **kwargs):

        self.mantencion = kwargs.pop(
            'mantencion',
            None
        )

        super().__init__(*args, **kwargs)

    def clean(self):

        cleaned_data = super().clean()

        fecha_cierre = cleaned_data.get(
            'fecha_cierre'
        )

        kilometraje_cierre = cleaned_data.get(
            'kilometraje_cierre'
        )

        if not self.mantencion:

            return cleaned_data

        if (
            self.mantencion.fecha_ingreso
            and fecha_cierre
            and fecha_cierre < self.mantencion.fecha_ingreso
        ):

            self.add_error(
                'fecha_cierre',
                'La fecha de cierre no puede ser anterior a la fecha de ingreso de la mantención.'
            )

        if (
            self.mantencion.tipo_mantencion == 'KILOMETRAJE'
            and kilometraje_cierre is None
        ):

            self.add_error(
                'kilometraje_cierre',
                'Debes indicar el kilometraje de cierre para una mantención por kilometraje.'
            )

        if (
            self.mantencion.kilometraje_ingreso is not None
            and kilometraje_cierre is not None
            and kilometraje_cierre < self.mantencion.kilometraje_ingreso
        ):

            self.add_error(
                'kilometraje_cierre',
                'El kilometraje de cierre no puede ser menor al kilometraje de ingreso.'
            )

        return cleaned_data
        
class CancelarMantencionVehiculoForm(forms.Form):

    observacion_cancelacion = forms.CharField(
        required=True,
        widget=forms.Textarea(
            attrs={
                'class': 'form-control',
                'rows': 4,
                'placeholder': 'Indica el motivo de cancelación de la mantención...'
            }
        ),
        label='Motivo de cancelación'
    )


class ReprogramarMantencionVehiculoForm(forms.Form):

    nueva_fecha_programada = forms.DateField(
        required=True,
        label='Nueva fecha programada',
        widget=forms.DateInput(
            attrs={
                'type': 'date'
            }
        )
    )

    motivo = forms.CharField(
        required=True,
        label='Motivo',
        max_length=255,
        widget=forms.TextInput(
            attrs={
                'placeholder': 'Motivo de la reprogramación'
            }
        )
    )

    observacion = forms.CharField(
        required=False,
        label='Observación',
        widget=forms.Textarea(
            attrs={
                'rows': 4,
                'placeholder': 'Observación de la reprogramación...'
            }
        )
    )

    def __init__(
        self,
        *args,
        **kwargs
    ):

        self.mantencion = kwargs.pop(
            'mantencion',
            None
        )

        super().__init__(
            *args,
            **kwargs
        )

    def clean(self):

        cleaned_data = super().clean()

        nueva_fecha_programada = cleaned_data.get(
            'nueva_fecha_programada'
        )

        if self.mantencion:

            if self.mantencion.estado == 'CERRADA':

                raise forms.ValidationError(
                    'No puedes reprogramar una mantención cerrada.'
                )

            if self.mantencion.estado == 'CANCELADA':

                raise forms.ValidationError(
                    'No puedes reprogramar una mantención cancelada.'
                )

            if self.mantencion.estado == 'EN_CURSO':

                raise forms.ValidationError(
                    'No puedes reprogramar una mantención que ya está en curso.'
                )

        if not nueva_fecha_programada:

            self.add_error(
                'nueva_fecha_programada',
                'Debes ingresar una nueva fecha programada.'
            )

        return cleaned_data
