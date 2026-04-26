from django.contrib import admin
from django import forms
from django.contrib.admin import SimpleListFilter
from .models import Evento, Documento, Recado, DocumentoPrivado, EventoPrivado, RecadoInterno, PeriodoAula, RegistroFalta
from ckeditor.widgets import CKEditorWidget

def admin_superuser_has_permission(request):
    return request.user.is_active and request.user.is_superuser


admin.site.has_permission = admin_superuser_has_permission


# ===== EVENTOS PÚBLICOS COM CKEDITOR =====
class EventoAdminForm(forms.ModelForm):
    descricao = forms.CharField(widget=CKEditorWidget())

    class Meta:
        model = Evento
        fields = '__all__'

class EventoAdmin(admin.ModelAdmin):
    form = EventoAdminForm
    fields = ['titulo', 'descricao', 'data', 'local', 'imagem', 'criado_por']
    list_display = ['titulo', 'data', 'local']

# ===== EVENTOS PRIVADOS COM CKEDITOR =====
class EventoPrivadoAdminForm(forms.ModelForm):
    descricao = forms.CharField(widget=CKEditorWidget())

    class Meta:
        model = EventoPrivado
        fields = '__all__'

class EventoPrivadoAdmin(admin.ModelAdmin):
    form = EventoPrivadoAdminForm
    fields = ['titulo', 'descricao', 'data_inicio', 'data_fim', 'local', 'criado_por']
    list_display = ['titulo', 'data_inicio', 'local']

# ===== RECADOS PÚBLICOS COM CKEDITOR =====
class RecadoAdminForm(forms.ModelForm):
    conteudo = forms.CharField(widget=CKEditorWidget())

    class Meta:
        model = Recado
        fields = '__all__'

class RecadoAdmin(admin.ModelAdmin):
    form = RecadoAdminForm
    list_display = ['titulo', 'fixado', 'criado_em']
    list_filter = ['fixado']

# ===== RECADOS INTERNOS COM CKEDITOR =====
class RecadoInternoAdminForm(forms.ModelForm):
    mensagem = forms.CharField(widget=CKEditorWidget())

    class Meta:
        model = RecadoInterno
        fields = '__all__'

class RecadoInternoAdmin(admin.ModelAdmin):
    form = RecadoInternoAdminForm
    list_display = ['id', 'criado_por', 'criado_em']
    list_filter = ['criado_em']

# ===== REGISTROS =====
admin.site.register(Evento, EventoAdmin)
admin.site.register(Documento)
admin.site.register(Recado, RecadoAdmin)
admin.site.register(DocumentoPrivado)
admin.site.register(EventoPrivado, EventoPrivadoAdmin)
admin.site.register(RecadoInterno, RecadoInternoAdmin)
admin.site.register(PeriodoAula)

# ===== VÍDEOS DO COLÉGIO =====
from .models import Video
admin.site.register(Video)

# ===== REGISTRO DE PONTO =====
from .models import RegistroPonto

# ===== CONTROLE DE FALTAS - ALUNOS =====
from .models import Turma, Aluno, RegistroFaltaAluno, DigitadorTurma

# ===== FILTRO PERSONALIZADO PARA TELEFONE =====
class TelefoneFilter(SimpleListFilter):
    title = 'Situação do Telefone'
    parameter_name = 'telefone_status'

    def lookups(self, request, model_admin):
        return (
            ('com', '✅ Com telefone'),
            ('sem', '❌ Sem telefone'),
        )

    def queryset(self, request, queryset):
        if self.value() == 'com':
            return queryset.exclude(telefone__isnull=True).exclude(telefone='')
        if self.value() == 'sem':
            return queryset.filter(telefone__isnull=True) | queryset.filter(telefone='')
        return queryset


@admin.register(Turma)
class TurmaAdmin(admin.ModelAdmin):
    list_display = ['nome', 'serie', 'ano', 'ativa']
    list_filter = ['ativa', 'ano']
    search_fields = ['nome', 'serie']

@admin.register(Aluno)
class AlunoAdmin(admin.ModelAdmin):
    list_display = ['nome', 'numero', 'turma', 'telefone', 'ativo']
    list_filter = ['turma', 'ativo', TelefoneFilter]
    search_fields = ['nome', 'numero']
    list_editable = ['ativo']


@admin.register(DigitadorTurma)
class DigitadorTurmaAdmin(admin.ModelAdmin):
    list_display = ['usuario', 'turma', 'ativo', 'criado_em']
    list_filter = ['ativo', 'turma']
    search_fields = ['usuario__username', 'usuario__first_name', 'usuario__last_name', 'turma__nome']

@admin.register(RegistroFaltaAluno)
class RegistroFaltaAlunoAdmin(admin.ModelAdmin):
    list_display = ['aluno', 'data', 'quantidade_faltas', 'justificada', 'responsavel_contatado']
    list_filter = ['justificada', 'data', 'aluno__turma']
    search_fields = ['aluno__nome', 'observacoes']
    date_hierarchy = 'data'
    raw_id_fields = ['aluno']

from import_export import resources
from import_export.admin import ImportExportMixin
from .models import Professor

class ProfessorResource(resources.ModelResource):
    class Meta:
        model = Professor
        fields = ['cpf', 'nome_completo', 'disciplinas', 'carga_horaria']
        import_id_fields = ['cpf']

class ProfessorAdmin(ImportExportMixin, admin.ModelAdmin):
    resource_class = ProfessorResource
    list_display = ('nome_completo', 'cpf', 'disciplinas_resumidas', 'carga_horaria', 'ativo')
    list_filter = ('ativo', 'cidade')
    search_fields = ('nome_completo', 'email', 'cpf', 'celular')
    list_editable = ('ativo',)

    fieldsets = (
        ('Dados Pessoais', {
            'fields': ('usuario', 'nome_completo', 'nome_social', 'data_nascimento', 'cpf', 'rg')
        }),
        ('Contato', {
            'fields': ('email', 'email_pessoal', 'telefone', 'celular')
        }),
        ('Endereço', {
            'fields': ('endereco', 'bairro', 'cidade', 'cep')
        }),
        ('Dados Profissionais', {
            'fields': ('disciplinas', 'formacao', 'data_admissao', 'carga_horaria')
        }),
        ('Documentos', {
            'fields': ('curriculo', 'documento_identidade', 'comprovante_residencia')
        }),
        ('Status', {
            'fields': ('ativo', 'observacoes')
        }),
    )

    def disciplinas_resumidas(self, obj):
        return obj.disciplinas[:50] + '...' if len(obj.disciplinas) > 50 else obj.disciplinas
    disciplinas_resumidas.short_description = 'Disciplinas'

admin.site.register(Professor, ProfessorAdmin)

# ===== REGISTRO DE OCORRÊNCIAS DE ALUNOS =====
from .models import RegistroOcorrenciaAluno

@admin.register(RegistroOcorrenciaAluno)
class RegistroOcorrenciaAlunoAdmin(admin.ModelAdmin):
    list_display = [
        'id', 'aluno', 'data', 'tipo_ocorrencia', 'turno',
        'faltou', 'busca_ativa_realizada',
        'resposta_disparo', 'resposta_disparo_data'  # ← ADICIONAR
    ]
    list_filter = ['data', 'tipo_ocorrencia', 'turno', 'busca_ativa_realizada', 'aluno__turma']
    search_fields = ['aluno__nome', 'aluno__turma__nome', 'motivo_alegado', 'resposta_disparo']
    list_per_page = 50
    date_hierarchy = 'data'
    raw_id_fields = ['aluno', 'registrado_por']
