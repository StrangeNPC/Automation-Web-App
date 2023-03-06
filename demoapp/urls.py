from django.urls import path 
from . import views 

urlpatterns = [ 
    path('', views.index, name='index'),
    path('upload-files/', views.upload_files, name='upload_files'),
]