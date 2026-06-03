def permisos_usuario(request):

    user = request.user

    puede_visualizar = False
    puede_editar = False
    puede_importar = False
    puede_gestionar_usuarios = False

    if user.is_authenticated:

        es_master = user.groups.filter(
            name='Master'
        ).exists()

        es_editor = user.groups.filter(
            name='Editor'
        ).exists()

        es_visualizador = user.groups.filter(
            name='Visualizador'
        ).exists()

        puede_visualizar = (
            es_visualizador
            or es_editor
            or es_master
        )

        puede_editar = (
            es_editor
            or es_master
        )

        puede_importar = es_master

        puede_gestionar_usuarios = es_master

    return {
        'puede_visualizar': puede_visualizar,
        'puede_editar': puede_editar,
        'puede_importar': puede_importar,
        'puede_gestionar_usuarios': puede_gestionar_usuarios,
    }