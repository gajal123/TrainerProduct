# -*- coding: utf-8 -*-
# Generated by Django 1.9 on 2017-07-25 13:14
from __future__ import unicode_literals

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('trainer', '0008_curriculum'),
    ]

    operations = [
        migrations.AddField(
            model_name='trainer',
            name='email',
            field=models.CharField(default='@gmail.com', max_length=200),
        ),
    ]
