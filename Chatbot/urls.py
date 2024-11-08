from django.urls import path
from . import views

app_name = "Chatbot"

urlpatterns = [
    path('accueil_chatbot', views.accueil_chatbot, name='accueil_chatbot'),
    path('chatbot_app', views.chatbot_application, name='chatbot_app'),
    path('documents_chatbot', views.documents_chatbot, name='documents_chatbot'),
    path('ajout_doc_temp', views.ajout_document_template, name='ajout_doc_temp'),
    path('ajout_doc_fct', views.ajout_document_fonction, name='ajout_doc_fct'),
    path('remove_doc/<str:document_name>', views.remove_doc, name='remove_doc'), 
    path('get_sources', views.get_sources, name='get_sources'),
    path('documents_chatbot2', views.documents_chatbot2, name='documents_chatbot2'),
    

]
