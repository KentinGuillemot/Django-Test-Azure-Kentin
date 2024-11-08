from django.contrib import admin
from .models import Bulletin, Bulletin_Net_Vers_Brut, Net_to_Brut, Bulletin_Cout_Vers_Brut, Cout_to_Brut

@admin.register(Bulletin)
class BulletinAdmin(admin.ModelAdmin):
    list_display = ('user','heures_semaine', 'heures_mois', 'type_salarie', 'salaire_brut_heure', 'salaire_brut_mois')
    search_fields = ('type_salarie',)
    list_filter = ('prevoyance', 'mutuelle')

@admin.register(Bulletin_Net_Vers_Brut)
class BulletinNetVersBrutAdmin(admin.ModelAdmin):
    list_display = ('user','heures_semaine', 'heures_mois', 'type_salarie', 'salaire_net_heure', 'salaire_net_mois')
    search_fields = ('type_salarie',)
    list_filter = ('prevoyance', 'mutuelle')

@admin.register(Net_to_Brut)
class NetToBrutAdmin(admin.ModelAdmin):
    list_display = ('user','salaire_brut', 'salaire_net_mois', 'bulletin')
    search_fields = ('bulletin__type_salarie',)

@admin.register(Bulletin_Cout_Vers_Brut)
class BulletinCoutVersBrutAdmin(admin.ModelAdmin):
    list_display = ('user','heures_semaine', 'heures_mois', 'type_salarie', 'cout_mois', 'cout_annee')
    search_fields = ('type_salarie',)
    list_filter = ('prevoyance', 'mutuelle')

@admin.register(Cout_to_Brut)
class CoutToBrutAdmin(admin.ModelAdmin):
    list_display = ('user','salaire_brut', 'cout_mois', 'bulletin')
    search_fields = ('bulletin__type_salarie',)
