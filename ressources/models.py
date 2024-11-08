from django.db import models
import json

class Condition(models.Model):
    nom = models.CharField(max_length=255)

    def __str__(self):
        return self.nom


class utilisateur(models.Model):
    nom = models.CharField(max_length=64, unique = False,verbose_name = "Nom")
    prenom = models.CharField(max_length=64, unique = False,verbose_name = "Prenom")
    email = models.CharField(max_length=64, unique = True,verbose_name = "Email")
    conditions_utili = models.CharField(max_length=200, unique = False)
    conditions_utili_m2m = models.ManyToManyField(Condition, blank=True)  # Nouveau champ

    def conditions_as_string(self):
        return ', '.join([condition.nom for condition in self.conditions_utili_m2m.all()])


    def set_conditions(self, lst):
        self.condtions_utili = json.dumps(lst)

    def get_conditions(self):
        return json.loads(self.condtions_utili)

    def __str__(self):
         return f"{self.prenom} {self.nom}"

    class Meta:
         verbose_name = "Utilisateur"
         verbose_name_plural = "Utilisateurs"



class objet(models.Model):
    nom = models.CharField(max_length=64, unique = True,verbose_name = "nom")
    famille1 = models.CharField(max_length=64, null=True, blank=True, unique = False, verbose_name = "Famille1")
    famille2 = models.CharField(max_length=64, null=True, blank=True, unique = False, verbose_name = "Famille2")
    famille3 = models.CharField(max_length=64, null=True, blank=True, unique = False, verbose_name = "Famille3")
    famille4 = models.CharField(max_length=64, null=True, blank=True, unique = False, verbose_name = "Famille3")
    conditions_objet = models.CharField(max_length=200, unique = False)
    conditions_objet_m2m = models.ManyToManyField(Condition, blank=True)  # Nouveau champ

    def conditions_as_string(self):
        return ', '.join([condition.nom for condition in self.conditions_objet_m2m.all()])


    def __str__(self):
         return self.nom
    
    def get_famille1_color(self):
        famille1_colors = {
            'Informatique': '#C8AD7F',
            'Véhicule': '#597535',
            # Ajoutez plus de familles et couleurs selon vos besoins
        }
        return famille1_colors.get(self.famille1, '#000000')  # Noir par défaut

    class Meta:
         verbose_name = "Objet"
         verbose_name_plural = "Objets"   


class resa(models.Model):
    obj = models.ForeignKey(objet, on_delete=models.CASCADE, verbose_name="Objet")
    user = models.ForeignKey(utilisateur, on_delete=models.CASCADE, verbose_name="Utilisateur")
    date_debut = models.DateTimeField()
    date_fin = models.DateTimeField()
    
    def __str__(self):
        return f"Réservation de {self.obj.nom} par {self.user}"

    class Meta:
        verbose_name = "Reservation"
        verbose_name_plural = "Reservations"

    def check_conflict(self):
        conflicts = resa.objects.filter(
            obj=self.obj,
            date_debut__lt=self.date_fin,
            date_fin__gt=self.date_debut
        )
        return conflicts
 


class nom_famille(models.Model):
    famille1 = models.CharField(max_length=64, default="Famille1", null=True, blank=True, unique = False, verbose_name = "Famille1")
    famille2 = models.CharField(max_length=64, default="Famille2", null=True, blank=True, unique = False, verbose_name = "Famille2")
    famille3 = models.CharField(max_length=64, default="Famille3", null=True, blank=True, unique = False, verbose_name = "Famille3")
    famille4 = models.CharField(max_length=64, default="Famille4", null=True, blank=True, unique = False, verbose_name = "Famille3")

    def __str__(self):
         return f"Famille1 : {self.famille1}, Famille2 : {self.famille2}, Famille2 : {self.famille2}, Famille2 : {self.famille2},"

    class Meta:
         verbose_name = "Objet"
         verbose_name_plural = "Objets"  