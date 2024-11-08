# accounts/admin.py
from django.contrib import admin
from .models import Entreprise, Salarie



@admin.action(description="Valider les entreprises clientes")
def validate_clients(modeladmin, request, queryset):
    queryset.update(validated=True)



class EntrepriseAdmin(admin.ModelAdmin):
    list_display = ('nom_entreprise', 'email', 'est_client','validated')
    list_filter = ('est_client', 'nom_entreprise','validated')  # Filtre sur le statut de client et le nom de l'entreprise
    search_fields = ('nom_entreprise', 'user__email')  # Recherche sur le nom de l'entreprise et l'e-mail associé
    actions = [validate_clients]  # Ajoute l'action de validation des entreprises clientes

    def email(self, obj):
        return obj.user.email
    email.short_description = 'Adresse e-mail'


class SalarieAdmin(admin.ModelAdmin):
    list_display = ('prenom', 'nom', 'email', 'entreprise')
    list_filter = ('entreprise',)  # Filtre les salariés par entreprise
    search_fields = ('prenom', 'nom', 'user__email', 'entreprise__nom_entreprise')  # Recherche sur prénom, nom, e-mail et entreprise

    def email(self, obj):
        return obj.user.email
    email.short_description = 'Adresse e-mail'


admin.site.register(Entreprise, EntrepriseAdmin)
admin.site.register(Salarie, SalarieAdmin)
