from django.urls import path
from . import views

urlpatterns = [
    path('read_excel/', views.read_excel, name='read_excel'),
    path('update_excel/', views.update_excel, name='update_excel'),
]