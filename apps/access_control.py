from functools import wraps

from django.contrib import messages
from django.shortcuts import redirect
from django.contrib.auth.decorators import user_passes_test


ADMIN_GROUPS = {'Administrador'}
EQUIPE_GROUPS = {'Equipe pedagógica', 'Equipe Pedagógica', 'Equipe Diretiva'}
PROFESSOR_GROUPS = {'Professor', 'Professores'}
DIGITADOR_GROUPS = {'Aluno digitador', 'Aluno Digitador', 'Digitadores'}


def user_in_groups(user, names):
    if not user.is_authenticated:
        return False
    return user.groups.filter(name__in=names).exists()


def user_in_groups_normalized(user, names):
    if not user.is_authenticated:
        return False
    normalized_names = {(name or '').strip().casefold() for name in names}
    user_group_names = user.groups.values_list('name', flat=True)
    return any((name or '').strip().casefold() in normalized_names for name in user_group_names)


def is_administrador(user):
    return user.is_authenticated and (user.is_superuser or user_in_groups(user, ADMIN_GROUPS))


def is_equipe_pedagogica(user):
    return is_administrador(user) or user_in_groups(user, EQUIPE_GROUPS)


def is_professor(user):
    return is_administrador(user) or user_in_groups_normalized(user, PROFESSOR_GROUPS)


def is_aluno_digitador(user):
    return user.is_authenticated and user_in_groups(user, DIGITADOR_GROUPS)


def get_login_redirect_url(user):
    if is_administrador(user):
        return '/inicio/'
    if is_equipe_pedagogica(user):
        return '/painel-equipe/'
    if is_professor(user):
        return '/inicio/'
    if is_aluno_digitador(user):
        return '/digitador/'
    return '/inicio/'


def can_access_painel(user):
    return is_equipe_pedagogica(user)


def can_manage_faltas_ocorrencias_relatorios(user):
    return is_equipe_pedagogica(user)


def can_register_professor_falta(user):
    return is_equipe_pedagogica(user) or is_professor(user)


def can_register_aluno_falta(user):
    return is_equipe_pedagogica(user) or is_professor(user) or is_aluno_digitador(user)


def can_register_ocorrencia(user):
    return is_equipe_pedagogica(user) or is_professor(user)


def can_consult_alunos(user):
    return is_equipe_pedagogica(user) or is_professor(user)


def permission_required(test_func, login_url='/'):
    return user_passes_test(test_func, login_url=login_url)


def superuser_required(view_func):
    @wraps(view_func)
    def wrapper(request, *args, **kwargs):
        if request.user.is_authenticated and request.user.is_superuser:
            return view_func(request, *args, **kwargs)
        messages.error(request, 'Acesso negado. Apenas superusuário.')
        return redirect('painel_equipe' if request.user.is_authenticated else 'home')

    return wrapper
