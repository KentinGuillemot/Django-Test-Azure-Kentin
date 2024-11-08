from django.shortcuts import render, get_object_or_404, redirect
from django.http import HttpResponse, JsonResponse #juste pour renvoyer simpliment un message
from .models import utilisateur, objet, resa  #j’import mes classes de modèles
import datetime
from django.contrib import messages
from django.core.serializers import serialize
from .forms import utili_form, ObjetForm, ResaForm,Famille1Form,Famille2Form,Famille3Form,Famille4Form,ConditionForm
import json

from django.shortcuts import render
from .models import utilisateur, objet, resa, Condition
from django.db.models import Count
import random
import calendar


def accueil(request):
    return render(request,"ressources/accueil.html",{})


def index(request):
    # Fetch statistics
    total_users = utilisateur.objects.count()
    total_objects = objet.objects.count()
    total_reservations = resa.objects.count()

    # Example statistics
    statistics = [
        {"title": "Utilisateurs", "value": total_users, "color": "primary"},
        {"title": "Objets", "value": total_objects, "color": "success"},
        {"title": "Réservations", "value": total_reservations, "color": "info"},
    ]

    latest_users = utilisateur.objects.all().order_by('-id')[:4]
    latest_objects = objet.objects.all().order_by('-id')[:4]
    latest_reservations = resa.objects.all().order_by('-id')[:4]

    # Fetch random data for tabs
    random_users = utilisateur.objects.order_by('?')[:4]
    random_objects = objet.objects.order_by('?')[:4]
    random_reservations = resa.objects.order_by('?')[:4]

    # Prepare data for reservations chart
    reservation_data = resa.objects.values('date_debut__month').annotate(count=Count('id')).order_by('date_debut__month')
    reservation_months = [calendar.month_abbr[item['date_debut__month']] for item in reservation_data]
    reservation_counts = [item['count'] for item in reservation_data]

    context = {
        'statistics': statistics,
        'latest_users': latest_users,
        'latest_objects': latest_objects,
        'latest_reservations': latest_reservations,
        'random_users': random_users,
        'random_objects': random_objects,
        'random_reservations': random_reservations,
        'reservation_months': reservation_months,
        'reservation_counts': reservation_counts,
    }
    return render(request, 'ressources/index.html', context)


def liste_utili(request):
    utilisateurs = utilisateur.objects.all()
    context = {
        "message": datetime.datetime.now(),
        "utilisateurs": utilisateurs,
        "objets": objet.objects.all(),
        "reservations": resa.objects.all()
    }
    return render(request, "ressources/utili-liste.html", context)

def liste_objet(request):
    objets = objet.objects.all()
    context = {
        "message": datetime.datetime.now(),
        "utilisateurs": utilisateur.objects.all(),
        "objets": objets,
        "reservations": resa.objects.all()
    }
    return render(request, "ressources/objet-liste.html", context)

def liste_resa(request):
    reservations = resa.objects.all()
    context = {
        "message": datetime.datetime.now(),
        "utilisateurs": utilisateur.objects.all(),
        "objets": objet.objects.all(),
        "reservations": reservations
    }
    return render(request, "ressources/resa-liste.html", context)

def liste_condition(request): 
    context = {"message" : datetime.datetime.now(),
               "utilisateurs" : utilisateur.objects.all(),
               "objets" : objet.objects.all(),
               "reservations" : resa.objects.all(),
               "conditions" : Condition.objects.all()
               }
    return render(request,"ressources/condition-liste.html",context)




def create_condition(request):
    if request.method == 'POST':
        form = ConditionForm(request.POST)
        if form.is_valid():
            form.save()
            return redirect('ressources:index')
    else:
        form = ConditionForm()
    return render(request, 'ressources/condition_form.html', {'form': form})

def create_user(request):
    if request.method == "POST":
        form = utili_form(request.POST)
        if form.is_valid():
            form.save()
            return redirect("ressources:utili-liste")
    else:
        form = utili_form()

    return render(request,"ressources/utili-form.html",{"form" : form})   

def create_objet(request):
    if request.method == 'POST':
        form = ObjetForm(request.POST)

        if form.is_valid():
            form.save()
            return redirect("ressources:objet-liste")  # Redirigez vers une page de succès ou ailleurs
    else:
        form = ObjetForm()

    return render(request, "ressources/objet-form.html", {'form': form})   



def create_resa(request, obj_id):
    obj = get_object_or_404(objet, pk=obj_id)
    if request.method == 'POST':
        form = ResaForm(request.POST)
        if form.is_valid():
            form.save()
            return redirect('ressources:resa-liste')
    else:
        obj_conditions = set(obj.conditions_objet_m2m.values_list('id', flat=True))
        valid_users = utilisateur.objects.all()
        filtered_users = []
        for user in valid_users:
            user_conditions = set(user.conditions_utili_m2m.values_list('id', flat=True))
            if obj_conditions.issubset(user_conditions):
                filtered_users.append(user)
        
        form = ResaForm(initial={'obj': obj}, user_queryset=utilisateur.objects.filter(id__in=[user.id for user in filtered_users]))
    
    return render(request, 'ressources/resa-form.html', {'form': form, 'objets': obj})

def create_resa2(request): 
    context = {"message" : datetime.datetime.now(),
               "utilisateurs" : utilisateur.objects.all(),
               "objets" : objet.objects.all(),
               "reservations" : resa.objects.all()
               }
    return render(request,"ressources/resa-crea.html",context)


def edit_user(request,utili_id):
    book = utilisateur.objects.get(pk = utili_id)

    if request.method == "POST": #le formulaire utilise la validation POST
        form = utili_form(request.POST, instance=book) #on lui dis qu'on a une instance pour pas avoir un formulaire avec des infos vide (vu qu'on veut modifier)

        if form.is_valid():
            form.save()
            return redirect("ressources:index")
    else:
        form = utili_form(instance=book)

    return render(request,"ressources/utili-edit.html",{"form" : form})

def edit_objet(request,obj_id):
    book = objet.objects.get(pk = obj_id)

    if request.method == "POST": #le formulaire utilise la validation POST
        form = ObjetForm(request.POST, instance=book) #on lui dis qu'on a une instance pour pas avoir un formulaire avec des infos vide (vu qu'on veut modifier)

        if form.is_valid():
            form.save()
            return redirect("ressources:index")
    else:
        form = ObjetForm(instance=book)

    return render(request,"ressources/objet-edit.html",{"form" : form})

def edit_resa(request,resa_id):
    book = resa.objects.get(pk = resa_id)

    if request.method == "POST": #le formulaire utilise la validation POST
        form = ResaForm(request.POST, instance=book) #on lui dis qu'on a une instance pour pas avoir un formulaire avec des infos vide (vu qu'on veut modifier)

        if form.is_valid():
            form.save()
            return redirect("ressources:index")
    else:
        form = ResaForm(instance=book)

    return render(request,"ressources/resa-edit.html",{"form" : form})

def edit_condi(request,resa_id):
    book = Condition.objects.get(pk = resa_id)

    if request.method == "POST": #le formulaire utilise la validation POST
        form = ConditionForm(request.POST, instance=book) #on lui dis qu'on a une instance pour pas avoir un formulaire avec des infos vide (vu qu'on veut modifier)

        if form.is_valid():
            form.save()
            return redirect("ressources:index")
    else:
        form = ConditionForm(instance=book)

    return render(request,"ressources/condition-edit.html",{"form" : form})





def remove_user(request, utili_id):
    try:
        book = utilisateur.objects.get(pk = utili_id)
        book.delete()
    except utilisateur.DoesNotExist:
        return redirect("ressources:utili-liste")
    return redirect("ressources:utili-liste")

def remove_objet(request, obj_id):
    try:
        book = objet.objects.get(pk = obj_id)
        book.delete()
    except objet.DoesNotExist:
        return redirect("ressources:objet-liste")
    return redirect("ressources:objet-liste")

def remove_resa(request, resa_id):
    try:
        book = resa.objects.get(pk = resa_id)
        book.delete()
    except resa.DoesNotExist:
        return redirect("ressources:resa-liste")
    return redirect("ressources:resa-liste")

def remove_condi(request, condi_id):
    try:
        book = Condition.objects.get(pk = condi_id)
        book.delete()
    except Condition.DoesNotExist:
        return redirect("ressources:condi-liste")
    return redirect("ressources:condi-liste")





def calendar_events(request):
    objet_id = request.GET.get('objet_id')
    famille1 = request.GET.get('famille1')
    famille2 = request.GET.get('famille2')
    famille3 = request.GET.get('famille3')
    famille4 = request.GET.get('famille4')
    filters = {}
    if objet_id:
        filters['obj__id'] = objet_id
    if famille1:
        filters['obj__famille1'] = famille1
    if famille2:
        filters['obj__famille2'] = famille2
    if famille3:
        filters['obj__famille3'] = famille3
    if famille4:
        filters['obj__famille4'] = famille4
    reservations = resa.objects.filter(**filters)
    
    events = [
        {
            'title': f'{resa.obj.nom} - {resa.user.nom}',
            'start': resa.date_debut.strftime('%Y-%m-%dT%H:%M:%S'),
            'end': resa.date_fin.strftime('%Y-%m-%dT%H:%M:%S'),
            'backgroundColor': resa.obj.get_famille1_color(),
            'borderColor': resa.obj.get_famille1_color(),
        }
        for resa in reservations
    ]
    return JsonResponse(events, safe=False)

from django.template.defaultfilters import json_script


def calendar_view(request):
    objets = list(objet.objects.all())
    famille1 = list(objet.objects.order_by('famille1').values_list('famille1', flat=True).distinct())
    famille2 = list(objet.objects.order_by('famille2').values_list('famille2', flat=True).distinct())
    famille3 = list(objet.objects.order_by('famille3').values_list('famille3', flat=True).distinct())
    famille4 = list(objet.objects.order_by('famille4').values_list('famille4', flat=True).distinct())

    objets_json = serialize('json', objets)

    return render(request, 'ressources/calendrier.html', {
        'objets': objets,
        'famille1': famille1,
        'famille2': famille2,
        'famille3': famille3,
        'famille4': famille4,
        'objets_json': objets_json,
    })

def choix_famille(request, nbr_fam):
    if nbr_fam == 1:
        if request.method == 'POST':
            form = Famille1Form(request.POST)
            if form.is_valid():
                form.save()
                return redirect("ressources:index")
        else:
            form = Famille1Form()

        return render(request, "ressources/famille-form.html", {'form': form})
    elif nbr_fam == 2:
        if request.method == 'POST':
            form = Famille2Form(request.POST)
            if form.is_valid():
                form.save()
                return redirect("ressources:index")
        else:
            form = Famille2Form()

        return render(request, "ressources/famille-form.html", {'form': form})
    elif nbr_fam == 3:
        if request.method == 'POST':
            form = Famille3Form(request.POST)
            if form.is_valid():
                form.save()
                return redirect("ressources:index")
        else:
            form = Famille3Form()

        return render(request, "ressources/famille-form.html", {'form': form})
    else:
        if request.method == 'POST':
            form = Famille4Form(request.POST)
            if form.is_valid():
                form.save()
                return redirect("ressources:index")
        else:
            form = Famille4Form()

        return render(request, "ressources/famille-form.html", {'form': form})
