from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("casting_plan", "0006_btrcampaign"),
    ]

    operations = [
        migrations.AddField(
            model_name="btrcampaign",
            name="rolling_plan_file",
            field=models.CharField(blank=True, max_length=255),
        ),
        migrations.AddField(
            model_name="btrcampaign",
            name="rolling_plan_created_at",
            field=models.DateTimeField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name="btrcampaign",
            name="rolling_plan_rm_qty",
            field=models.FloatField(default=0),
        ),
        migrations.AddField(
            model_name="btrcampaign",
            name="rolling_plan_bloom_count",
            field=models.PositiveIntegerField(default=0),
        ),
        migrations.AddField(
            model_name="btrcampaign",
            name="rolling_plan_yield_pct",
            field=models.FloatField(default=0),
        ),
    ]
