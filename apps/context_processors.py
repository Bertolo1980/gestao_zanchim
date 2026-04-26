from .access_control import (
    is_administrador,
    is_equipe_pedagogica,
    is_professor,
    is_aluno_digitador,
    can_access_painel,
    can_manage_faltas_ocorrencias_relatorios,
    can_register_professor_falta,
    can_register_aluno_falta,
    can_register_ocorrencia,
    can_consult_alunos,
)


def access_flags(request):
    user = request.user
    return {
        'acesso_administrador': is_administrador(user),
        'acesso_equipe_pedagogica': is_equipe_pedagogica(user),
        'acesso_professor': is_professor(user),
        'acesso_aluno_digitador': is_aluno_digitador(user),
        'pode_acessar_painel': can_access_painel(user),
        'pode_gerenciar_faltas_ocorrencias_relatorios': can_manage_faltas_ocorrencias_relatorios(user),
        'pode_registrar_falta_professor': can_register_professor_falta(user),
        'pode_registrar_falta_aluno': can_register_aluno_falta(user),
        'pode_registrar_ocorrencia': can_register_ocorrencia(user),
        'pode_consultar_alunos': can_consult_alunos(user),
    }
