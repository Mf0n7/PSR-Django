from django.urls import path
from . import views


urlpatterns = [
    path('', views.index, name='index'),
    path('upload/', views.upload_docx, name='upload_docx'),
    path('generate/', views.generate_docx, name='generate_docx'),
]
