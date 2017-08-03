from django.db import models
from django.core.urlresolvers import reverse
import datetime

class Trainer(models.Model):
    technology = models.CharField(max_length=250)
    name= models.CharField(max_length=250)
    location= models.CharField(max_length=100)
    email = models.CharField(max_length=200, default="@gmail.com")
    contact = models.CharField(max_length=20, default="+91")
    trainer_profile= models.FileField()

    def get_absolute_url(self):
        return reverse('trainer:details',kwargs={'pk' :self.pk})

    def __str__(self):
        return self.name + '-' + self.technology

class Curriculum(models.Model):
    name= models.CharField(max_length=100)
    course_curriculum= models.FileField()

    def __str__(self):
        return self.name

class Data(models.Model):
    name= models.CharField(max_length=100)
    contact = models.CharField(max_length=100)
    email = models.CharField(max_length=100 , unique=True)
    technology = models.CharField(max_length=100)


    def __str__(self):
        return self.name

class Document(models.Model):
    document = models.FileField(upload_to='sheets/')

    def __str__(self):
        return self.document