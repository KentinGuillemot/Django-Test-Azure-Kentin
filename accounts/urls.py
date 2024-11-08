from django.urls import path
from . import views

app_name = "accounts"

urlpatterns = [
    path('', views.accueil_user, name='accueil'),
    path('contact/', views.contact, name='contact'),
    path('login/', views.login_user, name='login'),
    path('logout/', views.logout_user, name='logout'),
    path('register/', views.register_user, name='register'),
    path('signup/entreprise/', views.entreprise_signup, name='entreprise_signup'),
    path('add_salarie/', views.add_salarie, name='add_salarie'),
    path('salarie_list/', views.salarie_list, name='salarie_list'),
    path('delete_salarie/<int:salarie_id>/', views.delete_salarie, name='delete_salarie'),
]
