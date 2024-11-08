from django.urls import path
from . import views

app_name = "ressources"

urlpatterns = [
    path('', views.index, name='index'),
    path('accueil', views.accueil, name='accueil'),
    path('ajouter-utilisateur/', views.create_user, name='create_user'),
    path('ajouter-objet/', views.create_objet, name='create_objet'),
    path('ajouter-resa/<int:obj_id>/', views.create_resa, name='create_resa'),
    path('edit-utilisateur/<int:utili_id>/', views.edit_user, name='edit_user'),
    path('edit-objet/<int:obj_id>/', views.edit_objet, name='edit_objet'),
    path('edit-resa/<int:resa_id>/', views.edit_resa, name='edit_resa'),
    path('edit-condi/<int:resa_id>/', views.edit_condi, name='edit_condi'),
    path('supp-utilisateur/<int:utili_id>/', views.remove_user, name='remove_user'),
    path('supp-objet/<int:obj_id>/', views.remove_objet, name='remove_objet'),
    path('supp-resa/<int:resa_id>/', views.remove_resa, name='remove_resa'),
    path('supp-condi/<int:condi_id>/', views.remove_condi, name='remove_condi'),
    path('events/', views.calendar_events, name='calendar_events'),
    path('calendar/', views.calendar_view, name='calendar_view'),
    path('choix-famille/<int:nbr_fam>/', views.choix_famille, name='choix_famille'),
    path('utili-liste', views.liste_utili, name='utili-liste'),
    path('objet-liste', views.liste_objet, name='objet-liste'),
    path('resa-liste', views.liste_resa, name='resa-liste'),
    path('ajout-resa', views.create_resa2, name='ajout-resa'),
    path('ajouter-condition/', views.create_condition, name='create_condition'),
    path('condi-liste/', views.liste_condition, name='condi-liste')
]
