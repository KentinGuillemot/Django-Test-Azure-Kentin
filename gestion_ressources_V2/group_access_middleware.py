from django.contrib import messages
from django.urls import reverse
from django.contrib.auth.models import Group
from django.shortcuts import redirect
from django.http import HttpResponse

class GroupAccessMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        # Définir les groupes autorisés pour chaque application
        allowed_apps_by_group = {
            'chatbot': ['Group_LIAM'],   # Groupes autorisés pour l'application 'chatbot'
            'ress': ['Group_EDGAR'],  # Groupes autorisés pour l'application 'ressources'
            'simulateur_paie': ['Group_Simu'], # Groupes autorisés pour l'application 'simulateur_paie'
        }

        # Vérifier si l'utilisateur est connecté
        if request.user.is_authenticated:
            # Parcourir chaque application et ses groupes autorisés
            for app, allowed_groups in allowed_apps_by_group.items():
                # Vérifier si le chemin d'URL de la requête correspond à une application autorisée
                if f"/{app}/" in request.path.lower():
                    # Vérifier si l'utilisateur appartient à au moins un des groupes autorisés pour cette application
                    if not any(request.user.groups.filter(name=group).exists() for group in allowed_groups):
                        # Bloquer l'accès si l'utilisateur n'est pas dans un groupe autorisé pour cette application
                        return HttpResponse("Accès restreint à cet outil", status=403)

        # Continuer avec la requête si tout est OK
        response = self.get_response(request)
        return response
