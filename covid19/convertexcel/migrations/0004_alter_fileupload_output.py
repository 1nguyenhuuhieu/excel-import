# Generated by Django 3.2.9 on 2021-11-16 04:22

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('convertexcel', '0003_alter_fileupload_output'),
    ]

    operations = [
        migrations.AlterField(
            model_name='fileupload',
            name='output',
            field=models.URLField(blank=True),
        ),
    ]
