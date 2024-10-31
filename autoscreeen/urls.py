from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),  # Home view for managing screenshots and files
    path('download/', views.download, name='download'),  # Endpoint for downloading the last created file
]
