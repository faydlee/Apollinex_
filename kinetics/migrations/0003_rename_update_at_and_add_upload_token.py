from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("kinetics", "0002_measurement"),
    ]

    operations = [
        migrations.RenameField(
            model_name="experiment",
            old_name="update_at",
            new_name="updated_at",
        ),
        migrations.AddField(
            model_name="experiment",
            name="upload_token",
            field=models.UUIDField(
                blank=True,
                editable=False,
                null=True,
                unique=True,
            ),
        ),
    ]
