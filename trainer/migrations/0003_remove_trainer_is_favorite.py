# -*- coding: utf-8 -*-
# Generated by Django 1.9 on 2017-07-17 06:58
from __future__ import unicode_literals

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('trainer', '0002_trainer_is_favorite'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='trainer',
            name='is_favorite',
        ),
    ]