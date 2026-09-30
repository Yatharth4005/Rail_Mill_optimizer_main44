from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('casting_plan', '0012_railmill_masterdata'),
    ]

    operations = [
        migrations.AddField(
            model_name='railmillcalculationmaster',
            name='rm_calculated_length_m',
            field=models.FloatField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name='railmillcalculationmaster',
            name='maximum_cooling_bed_length_m',
            field=models.FloatField(blank=True, null=True),
        ),
    ]
