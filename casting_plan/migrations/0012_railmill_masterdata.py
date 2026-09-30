from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('casting_plan', '0011_btrcampaign_embossing_plan_file'),
    ]

    operations = [
        migrations.CreateModel(
            name='RailMillCalculationMaster',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('planning_mat', models.CharField(blank=True, max_length=100)),
                ('rolled_section', models.CharField(blank=True, max_length=150)),
                ('material', models.CharField(blank=True, max_length=150)),
                ('product', models.CharField(blank=True, max_length=80)),
                ('sec_wt', models.FloatField(blank=True, null=True)),
                ('ext_grade', models.CharField(blank=True, max_length=100)),
                ('grade_group', models.CharField(blank=True, max_length=80)),
                ('rm_grade', models.CharField(blank=True, max_length=100)),
                ('rm_size', models.CharField(blank=True, max_length=100)),
                ('bloom_length_m', models.FloatField(blank=True, null=True)),
                ('bloom_weight_kg_m', models.FloatField(blank=True, null=True)),
                ('total_rolled_length_m', models.FloatField(blank=True, null=True)),
                ('total_finish_length_m', models.FloatField(blank=True, null=True)),
                ('cut_pcs', models.CharField(blank=True, max_length=100)),
                ('end_cut_margin', models.FloatField(blank=True, null=True)),
                ('end_cut_margin_req', models.FloatField(blank=True, null=True)),
                ('old_furnace_min_len', models.FloatField(blank=True, null=True)),
                ('old_furnace_max_len', models.FloatField(blank=True, null=True)),
                ('new_furnace_min_len', models.FloatField(blank=True, null=True)),
                ('new_furnace_max_len', models.FloatField(blank=True, null=True)),
            ],
            options={
                'verbose_name_plural': 'Rail Mill Calculation Master',
            },
        ),
        migrations.CreateModel(
            name='RailMillChemistryMaster',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('sl_no', models.CharField(blank=True, max_length=40)),
                ('product', models.CharField(blank=True, max_length=100)),
                ('internal_grade', models.CharField(blank=True, max_length=100)),
                ('proposed_internal_grade', models.CharField(blank=True, max_length=100)),
                ('ext_grade', models.CharField(blank=True, max_length=200)),
                ('sections', models.CharField(blank=True, max_length=200)),
                ('range_type', models.CharField(blank=True, max_length=30)),
                ('c_pct', models.CharField(blank=True, max_length=40)),
                ('mn_pct', models.CharField(blank=True, max_length=40)),
                ('s_pct', models.CharField(blank=True, max_length=40)),
                ('p_pct', models.CharField(blank=True, max_length=40)),
                ('si_pct', models.CharField(blank=True, max_length=40)),
                ('al_pct', models.CharField(blank=True, max_length=40)),
                ('cu_pct', models.CharField(blank=True, max_length=40)),
                ('ni_pct', models.CharField(blank=True, max_length=40)),
                ('cr_pct', models.CharField(blank=True, max_length=40)),
                ('mo_pct', models.CharField(blank=True, max_length=40)),
                ('ti_pct', models.CharField(blank=True, max_length=40)),
                ('v_pct', models.CharField(blank=True, max_length=40)),
                ('sn_pct', models.CharField(blank=True, max_length=40)),
                ('nb_pct', models.CharField(blank=True, max_length=40)),
                ('b_pct', models.CharField(blank=True, max_length=40)),
                ('as_pct', models.CharField(blank=True, max_length=40)),
                ('ce', models.CharField(blank=True, max_length=40)),
                ('n2_ppm', models.CharField(blank=True, max_length=40)),
                ('o2_ppm', models.CharField(blank=True, max_length=40)),
                ('h2_ppm', models.CharField(blank=True, max_length=40)),
            ],
            options={
                'verbose_name_plural': 'Rail Mill Chemistry Master',
            },
        ),
        migrations.CreateModel(
            name='RailMillPlanningMaterial',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('planning_mat', models.CharField(max_length=100, unique=True)),
            ],
            options={
                'verbose_name_plural': 'Rail Mill Planning Materials',
            },
        ),
        migrations.CreateModel(
            name='RailMillRMSectionWeight',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('rm_sec', models.CharField(max_length=100)),
                ('weight_t_per_m', models.FloatField()),
            ],
            options={
                'verbose_name_plural': 'Rail Mill RM Section Weights',
            },
        ),
        migrations.CreateModel(
            name='RailMillSawTolerance',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('condition', models.CharField(max_length=50)),
                ('tolerance_m', models.FloatField()),
            ],
            options={
                'verbose_name_plural': 'Rail Mill Saw Length Tolerances',
            },
        ),
        migrations.CreateModel(
            name='RailMillSectionGradeStandardMap',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('sl_no', models.CharField(blank=True, max_length=40)),
                ('rolled_section', models.CharField(blank=True, max_length=150)),
                ('flange_thickness_mm', models.FloatField(blank=True, null=True)),
                ('product', models.CharField(blank=True, max_length=80)),
                ('std_e250_e275', models.CharField(blank=True, max_length=150)),
                ('std_e350_s355', models.CharField(blank=True, max_length=150)),
                ('std_e410_jis', models.CharField(blank=True, max_length=150)),
                ('std_s420n', models.CharField(blank=True, max_length=150)),
                ('std_s460_e450', models.CharField(blank=True, max_length=150)),
                ('std_s500', models.CharField(blank=True, max_length=150)),
                ('std_e550', models.CharField(blank=True, max_length=150)),
                ('std_s355nl', models.CharField(blank=True, max_length=150)),
                ('fire_resistance', models.CharField(blank=True, max_length=150)),
                ('is_3443', models.CharField(blank=True, max_length=150)),
                ('irs_t12', models.CharField(blank=True, max_length=150)),
                ('en_13674', models.CharField(blank=True, max_length=150)),
            ],
            options={
                'verbose_name_plural': 'Rail Mill Section Grade Standards',
            },
        ),
        migrations.CreateModel(
            name='RailMillSecWiseMaster',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('section', models.CharField(blank=True, max_length=100)),
                ('section_weight_desc', models.CharField(blank=True, max_length=150)),
                ('input_size', models.CharField(blank=True, max_length=100)),
                ('bloom_length_m', models.FloatField(blank=True, null=True)),
                ('bloom_weight_kg_m', models.FloatField(blank=True, null=True)),
                ('total_rolled_length_m', models.FloatField(blank=True, null=True)),
                ('total_finish_length_m', models.FloatField(blank=True, null=True)),
                ('cut_pcs', models.CharField(blank=True, max_length=150)),
            ],
            options={
                'verbose_name_plural': 'Rail Mill Section Wise Matrix',
            },
        ),
    ]
