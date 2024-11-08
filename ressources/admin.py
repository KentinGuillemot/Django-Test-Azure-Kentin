from django.contrib import admin

from .models import utilisateur, objet, resa 

@admin.register(utilisateur)
class utilisateurAdmin(admin.ModelAdmin):
    fieldsets = [
        ("Information personnelles", {"fields":["nom","prenom","email"]}),
        ("Atouts", {'fields' : ["conditions_utili"]})
    ]

    list_display = ("nom","prenom","email","conditions_utili")
    search_fields = ["nom","conditions_utili"]

@admin.register(objet)
class objetAdmin(admin.ModelAdmin):
    fieldsets = [
        ("Nom", {"fields":["nom"]}),
        ("Familles", {'fields' : ["famille1","famille2","famille3","famille4"]}),
        ("Conditions d'utilisations", {"fields":["conditions_objet"]})
    ]

    list_display = ("nom","famille1","famille2","famille3","famille4","conditions_objet")
    list_filter = ["nom"]
    search_fields = ["nom","famille1","famille2","famille3","famille4","conditions_objet"]

@admin.register(resa)
class resaAdmin(admin.ModelAdmin):
    fieldsets = [
        ("Objet", {"fields":["obj"]}),
        ("Utilisateur", {'fields' : ["user"]}),
        ("Dates", {"fields":["date_debut","date_fin"]})
    ]

    list_display = ("obj","user","date_debut","date_fin")
    list_filter = ["obj","user","date_debut","date_fin"]
    search_fields = ["obj","user","date_debut","date_fin"]


