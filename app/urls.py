from django.urls import path
from . import views

app_name = "geo_location"

urlpatterns = [
    path('', views.home ,name="home"),
]