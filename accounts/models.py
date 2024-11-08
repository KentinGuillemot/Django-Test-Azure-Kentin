# accounts/models.py
from django.contrib.auth.models import AbstractUser, UserManager
from django.db import models
from django.conf import settings

class CustomUserManager(UserManager):
    pass

class CustomUser(AbstractUser):
    is_entreprise = models.BooleanField(default=False)
    is_salarie = models.BooleanField(default=False)

class Entreprise(models.Model):
    user = models.OneToOneField(settings.AUTH_USER_MODEL, on_delete=models.CASCADE)
    nom_entreprise = models.CharField(max_length=255)
    est_client = models.BooleanField(default=False)  # A valider manuellement dans l'admin
    validated = models.BooleanField(default=False)  # Indicates if admin has validated


    def __str__(self):
        return self.nom_entreprise

class Salarie(models.Model):
    entreprise = models.ForeignKey(Entreprise, on_delete=models.CASCADE, related_name='salaries')
    user = models.OneToOneField(settings.AUTH_USER_MODEL, on_delete=models.CASCADE)
    prenom = models.CharField(max_length=100)
    nom = models.CharField(max_length=100)

    def __str__(self):
        return f"{self.prenom} {self.nom}"
