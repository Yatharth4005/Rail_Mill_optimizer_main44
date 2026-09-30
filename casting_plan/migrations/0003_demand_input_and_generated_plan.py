from django.db import migrations, models

class Migration(migrations.Migration):
    dependencies = [('casting_plan', '0002_notification_alter_order_options')]
    operations = [
        migrations.AddField(model_name='order', name='so_item', field=models.CharField(blank=True, max_length=80)),
        migrations.AddField(model_name='order', name='pgi_qty', field=models.FloatField(default=0)),
        migrations.AddField(model_name='order', name='balance_qty', field=models.FloatField(default=0)),
        migrations.AddField(model_name='order', name='eligible_stock', field=models.FloatField(default=0)),
        migrations.AddField(model_name='order', name='net_demand', field=models.FloatField(default=0)),
        migrations.AddField(model_name='order', name='section_weight', field=models.FloatField(default=0)),
        migrations.AddField(model_name='order', name='required_delivery_date', field=models.CharField(blank=True, max_length=40)),
        migrations.AddField(model_name='order', name='source_file', field=models.CharField(blank=True, max_length=200)),
        migrations.AddField(model_name='order', name='raw_data', field=models.JSONField(blank=True, default=dict)),
        migrations.AddField(model_name='order', name='is_selected', field=models.BooleanField(default=True)),
        migrations.CreateModel(name='GeneratedPlan', fields=[
            ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
            ('created_at', models.DateTimeField(auto_now_add=True)),
            ('name', models.CharField(max_length=160)),
            ('planning_date', models.DateField(blank=True, null=True)),
            ('selected_order_count', models.PositiveIntegerField(default=0)),
            ('demand_qty', models.FloatField(default=0)),
            ('rm_qty', models.FloatField(default=0)),
            ('yield_pct', models.FloatField(default=0)),
            ('output_file', models.CharField(blank=True, max_length=255)),
        ])
    ]
