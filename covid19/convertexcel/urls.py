from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name="index" ),
    path('download/<str:id>/', views.view, name="view" ),
]