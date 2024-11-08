from django.shortcuts import redirect
from django.urls import reverse
from django.conf import settings


# Ici c'est pour dire que pour accèder à n'importe quoi faut etre connecté

class LoginRequiredMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        # Exempter certaines URLs pour éviter une redirection infinie
        exempt_urls = [
            reverse('accounts:login'),  # Remplacez avec le nom de la vue de connexion
            reverse('accounts:register'),  # Exempter la page d'inscription
            reverse('admin:index'),  # Admin Django
            reverse('accounts:entreprise_signup'),
            # Ajoutez d'autres URLs à exempter ici
        ]

        # Rediriger si l'utilisateur n'est pas authentifié et l'URL n'est pas exemptée
        if not request.user.is_authenticated and request.path not in exempt_urls:
            return redirect(f"{reverse('accounts:login')}?next={request.path}")

        response = self.get_response(request)
        return response
