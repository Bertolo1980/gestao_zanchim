from django.db import migrations


def create_access_groups(apps, schema_editor):
    Group = apps.get_model('auth', 'Group')
    for name in [
        'Administrador',
        'Equipe pedagógica',
        'Professor',
        'Aluno digitador',
    ]:
        Group.objects.get_or_create(name=name)


class Migration(migrations.Migration):

    dependencies = [
        ('apps', '0021_digitadorturma'),
        ('auth', '0012_alter_user_first_name_max_length'),
    ]

    operations = [
        migrations.RunPython(create_access_groups, migrations.RunPython.noop),
    ]
