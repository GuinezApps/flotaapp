from django.contrib.auth.decorators import user_passes_test
from django.shortcuts import redirect


def usuario_es_master(user):
    return (
        user.is_authenticated
        and user.groups.filter(name='Master').exists()
    )


def usuario_es_editor(user):
    return (
        user.is_authenticated
        and (
            user.groups.filter(name='Editor').exists()
            or user.groups.filter(name='Master').exists()
        )
    )


def usuario_es_visualizador(user):

    return (
        user.is_authenticated
        and (
            user.groups.filter(name='Visualizador').exists()
            or user.groups.filter(name='Editor').exists()
            or user.groups.filter(name='Master').exists()
        )
    )


def editor_required(view_func):

    return user_passes_test(
        usuario_es_editor,
        login_url='/login/'
    )(view_func)


def master_required(view_func):

    return user_passes_test(
        usuario_es_master,
        login_url='/login/'
    )(view_func)