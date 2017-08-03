from django.conf.urls import url
from .import views

app_name='trainer'

urlpatterns=[
    #/trainer/
    url(r'^$', views.IndexView.as_view(),name='index'),

    #/trainer/<trainer_id>/
    url(r'^(?P<pk>[0-9]+)/$',views.DetailView.as_view(),name='details'),

    #/trainer/trainer/add
    url(r'trainer/add/$', views.TrainerCreate.as_view(), name='Trainer-add'),

    #/trainer/trainer/<album_id>
    url(r'trainer/(?P<pk>[0-9]+)/$', views.TrainerUpdate.as_view(), name='Trainer-update'),

    #/trainer/trainer/add
    url(r'trainer/(?P<pk>[0-9]+)/delete/$', views.TrainerDelete.as_view(), name='Trainer-delete'),

    #/trainer/search/q=' '
    url(r'^search/$', views.search, name='Search'),

    url(r'^download/$', views.download, name='Trainer-download'),

    url(r'^proposalform/$', views.proposalform, name='Proposalform'),
    url(r'^proposalform2/$', views.proposalform2, name='Proposalform2'),
    url(r'^proposal/$', views.proposal, name='Proposal'),

    url(r'^upload/$', views.upload, name='upload'),

    url(r'^certificate/$', views.certificate, name='certificate'),
    url(r'^certificategenerate/$', views.certificategenerate, name='certificategenerate'),
    url(r'^uploadcertificate/$', views.uploadcertificate, name='uploadcertificate'),
    url(r'^excelupload/$', views.excelupload, name='excelupload'),


]