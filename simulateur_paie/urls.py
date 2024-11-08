from django.urls import path
from . import views

app_name = "simulateur_paie"

urlpatterns = [
    path('', views.create_bulletin, name='simu'),
    path('Accueil_simu', views.accueil, name='accueil'),
    path('telecharger_tableau/<int:bulletin_id>/', views.telecharger_tableau_pdf, name='telecharger_tableau'),
    path('historique', views.Historique,name='historique'),
    path('historique_net_to_brut', views.Historique_net_to_brut,name='historique_net_to_brut'),
    path('historique_cout_to_brut', views.Historique_cout_to_brut,name='historique_cout_to_brut'),
    path('bulletin/<int:bulletin_id>/', views.bulletin_detail, name='bulletin_detail'),
    path('edit_bulletin/<int:bulletin_id>/', views.edit_bulletin, name='edit_bulletin'),
    path('remove_bulletin/<int:bulletin_id>/', views.remove_bulletin, name='remove_bulletin'),
    path('bulletin2/<int:bulletin_id>/', views.bulletin2, name='bulletin2'),
    path('Net_to_Brut/', views.create_bulletin_Net_vers_Brut, name='Net_to_Brut'),
    path('Net_to_Brut_calcul/<int:bulletin_id>/', views.calculer_net_vers_brut, name='Net_to_Brut_calcul'),
    path('bulletin_net_to_brut/<int:bulletin_id>/', views.bulletin_detail_net_to_brut, name='bulletin_net_to_brut'), #le bulletin detail pour le net vers brut
    path('remove_bulletin_net_to_brut/<int:bulletin_id>/', views.remove_bulletin_net_to_brut, name='remove_bulletin_net_to_brut'),
    path('edit_bulletin_net_to_brut/<int:bulletin_id>/', views.edit_bulletin_net_to_brut, name='edit_bulletin_net_to_brut'),
    path('Cout_to_Brut/', views.create_bulletin_Cout_vers_Brut, name='Cout_to_Brut'),
    path('Cout_to_Brut_calcul/<int:bulletin_id>/', views.calculer_cout_vers_brut, name='Cout_to_Brut_calcul'),
    path('remove_bulletin_cout_to_brut/<int:bulletin_id>/', views.remove_bulletin_cout_to_brut, name='remove_bulletin_cout_to_brut'),
    path('bulletin_cout_to_brut/<int:bulletin_id>/', views.bulletin_detail_cout_to_brut, name='bulletin_cout_to_brut'), #le bulletin detail pour le cout vers brut
    path('edit_bulletin_cout_to_brut/<int:bulletin_id>/', views.edit_bulletin_cout_to_brut, name='edit_bulletin_cout_to_brut'),
    
]
