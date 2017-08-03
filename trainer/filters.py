from .models import Trainer
import django_filters

class UserFilter(django_filters.FilterSet):
    class Meta:
        model = Trainer
        fields = ['name', 'technology', 'location', ]