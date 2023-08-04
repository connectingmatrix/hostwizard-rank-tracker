from django.urls import path
from . import views


app_name = 'airbnb'  # here for namespacing of urls.

urlpatterns = [
    path("run_api/", views.run_bot, name="run_api"),
    path("homepage/", views.homepage, name="homepage"),
]