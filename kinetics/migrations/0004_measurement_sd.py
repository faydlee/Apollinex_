from django.db import migrations, models


class Migration(migrations.Migration):
    dependencies = [
        ("kinetics", "0003_rename_update_at_and_add_upload_token"),
    ]

    operations = [
        migrations.AddField(
            model_name="measurement",
            name="sd",
            field=models.FloatField(blank=True, default=None, null=True),
        ),
    ]
