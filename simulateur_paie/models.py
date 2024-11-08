from django.db import models
from django.conf import settings
from django.contrib.auth.models import User

class Bulletin(models.Model):

    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, null=True)

    heures_semaine = models.IntegerField() 
    heures_mois = models.IntegerField() 
    type_salarie = models.CharField(max_length=20) 
    prevoyance = models.BooleanField(default=False)
    type_plafond = models.CharField(max_length=20, null=True, blank=True)
    taux_prev_T1_sal = models.FloatField(default=0)
    taux_prev_T1_patr = models.FloatField(default=0)
    taux_prev_T2_sal = models.FloatField(default=0)
    taux_prev_T2_patr = models.FloatField(default=0)
    mutuelle = models.BooleanField(default=False)
    montant_mutu_sal = models.FloatField(default=0)
    montant_mutu_patr = models.FloatField(default=0)
    Taux_ATHT = models.FloatField(default=1.00)
    taux_ccss_red = models.CharField(max_length=20, null=True, blank=True)
    admin_SAM = models.CharField(max_length=20, null=True, blank=True)
    Exclu_ass_cho = models.CharField(max_length=20, null=True, blank=True)
    CCPB_ouvrier = models.CharField(max_length=20, null=True, blank=True)
    CCPB_Etam_cadre = models.CharField(max_length=20, null=True, blank=True)
    Gens_de_Maison = models.CharField(max_length=20, null=True, blank=True)
    salaire_brut_heure = models.FloatField()
    salaire_brut_mois = models.FloatField()
    remboursement_transport = models.FloatField(default=0)
    ticket_resto_prix = models.FloatField(default=0)
    ticket_resto_sal = models.FloatField(default=0)
    ticket_resto_patr = models.FloatField(default=0)
    panier_prix = models.FloatField(default=0)
    paniers_sal = models.FloatField(default=0)

class Bulletin_Net_Vers_Brut(models.Model):

    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, null=True)

    heures_semaine = models.IntegerField() 
    heures_mois = models.IntegerField() 
    type_salarie = models.CharField(max_length=20) 
    prevoyance = models.BooleanField(default=False)
    type_plafond = models.CharField(max_length=20, null=True, blank=True)
    taux_prev_T1_sal = models.FloatField(default=0)
    taux_prev_T1_patr = models.FloatField(default=0)
    taux_prev_T2_sal = models.FloatField(default=0)
    taux_prev_T2_patr = models.FloatField(default=0)
    mutuelle = models.BooleanField(default=False)
    montant_mutu_sal = models.FloatField(default=0)
    montant_mutu_patr = models.FloatField(default=0)
    Taux_ATHT = models.FloatField(default=1.00)
    taux_ccss_red = models.CharField(max_length=20, null=True, blank=True)
    admin_SAM = models.CharField(max_length=20, null=True, blank=True)
    Exclu_ass_cho = models.CharField(max_length=20, null=True, blank=True)
    CCPB_ouvrier = models.CharField(max_length=20, null=True, blank=True)
    CCPB_Etam_cadre = models.CharField(max_length=20, null=True, blank=True)
    Gens_de_Maison = models.CharField(max_length=20, null=True, blank=True)

    salaire_net_heure = models.FloatField()
    salaire_net_mois = models.FloatField()
    indemnite = models.CharField(max_length=20,null=True, blank=True) 

    remboursement_transport = models.FloatField(default=0)
    ticket_resto_prix = models.FloatField(default=0)
    ticket_resto_sal = models.FloatField(default=0)
    ticket_resto_patr = models.FloatField(default=0)
    panier_prix = models.FloatField(default=0)
    paniers_sal = models.FloatField(default=0)

class Net_to_Brut(models.Model):

    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, null=True)

    salaire_brut = models.FloatField()
    salaire_net_mois = models.FloatField(null=True, blank=True)
    bulletin = models.ForeignKey(Bulletin_Net_Vers_Brut, on_delete=models.CASCADE, null=True, blank=True)


class Bulletin_Cout_Vers_Brut(models.Model):

    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, null=True)

    heures_semaine = models.IntegerField() 
    heures_mois = models.IntegerField() 
    type_salarie = models.CharField(max_length=20) 
    prevoyance = models.BooleanField(default=False)
    type_plafond = models.CharField(max_length=20, null=True, blank=True)
    taux_prev_T1_sal = models.FloatField(default=0)
    taux_prev_T1_patr = models.FloatField(default=0)
    taux_prev_T2_sal = models.FloatField(default=0)
    taux_prev_T2_patr = models.FloatField(default=0)
    mutuelle = models.BooleanField(default=False)
    montant_mutu_sal = models.FloatField(default=0)
    montant_mutu_patr = models.FloatField(default=0)
    Taux_ATHT = models.FloatField(default=1.00)
    taux_ccss_red = models.CharField(max_length=20, null=True, blank=True)
    admin_SAM = models.CharField(max_length=20, null=True, blank=True)
    Exclu_ass_cho = models.CharField(max_length=20, null=True, blank=True)
    CCPB_ouvrier = models.CharField(max_length=20, null=True, blank=True)
    CCPB_Etam_cadre = models.CharField(max_length=20, null=True, blank=True)
    Gens_de_Maison = models.CharField(max_length=20, null=True, blank=True)

    cout_mois = models.FloatField()
    cout_annee = models.FloatField()
    indemnite = models.CharField(max_length=20,null=True, blank=True) 

    remboursement_transport = models.FloatField(default=0)
    ticket_resto_prix = models.FloatField(default=0)
    ticket_resto_sal = models.FloatField(default=0)
    ticket_resto_patr = models.FloatField(default=0)
    panier_prix = models.FloatField(default=0)
    paniers_sal = models.FloatField(default=0)

class Cout_to_Brut(models.Model):

    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, null=True)

    salaire_brut = models.FloatField()
    cout_mois = models.FloatField(null=True, blank=True)
    bulletin = models.ForeignKey(Bulletin_Cout_Vers_Brut, on_delete=models.CASCADE, null=True, blank=True)
