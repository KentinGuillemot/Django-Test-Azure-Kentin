# accounts/forms.py
from django import forms
from django.contrib.auth.forms import UserCreationForm
from .models import CustomUser, Entreprise, Salarie

# forms.py
class EntrepriseSignUpForm(UserCreationForm):
    nom_entreprise = forms.CharField(max_length=255)
    est_client = forms.BooleanField(required=False)

    class Meta:
        model = CustomUser
        fields = ('username', 'email', 'password1', 'password2')

    def save(self, commit=True):
        user = super().save(commit=False)
        user.is_entreprise = True
        if commit:
            user.save()
            Entreprise.objects.create(
                user=user,
                nom_entreprise=self.cleaned_data['nom_entreprise'],
                est_client=self.cleaned_data['est_client'],
                validated=False if self.cleaned_data['est_client'] else True
            )
        return user

class SalarieSignUpForm(UserCreationForm):
    prenom = forms.CharField(max_length=100)
    nom = forms.CharField(max_length=100)

    class Meta:
        model = CustomUser
        fields = ('username', 'email', 'password1', 'password2')

    def __init__(self, *args, **kwargs):
        self.entreprise = kwargs.pop('entreprise', None)
        super().__init__(*args, **kwargs)

    def save(self, commit=True):
        user = super().save(commit=False)
        user.is_salarie = True
        if commit:
            user.save()
            Salarie.objects.create(user=user, entreprise=self.entreprise,
                                   prenom=self.cleaned_data['prenom'],
                                   nom=self.cleaned_data['nom'])
        return user
