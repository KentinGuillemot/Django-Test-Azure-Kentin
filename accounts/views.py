from django.shortcuts import render, redirect
from django.contrib.auth.forms import AuthenticationForm, UserCreationForm
from django.contrib.auth import authenticate, login, logout
from django.contrib import messages
from django.contrib.auth.decorators import user_passes_test

import os
import pandas as pd
from django.conf import settings

from django.shortcuts import get_object_or_404
from django.contrib.auth.decorators import login_required
from django.http import HttpResponseRedirect
from django.urls import reverse
from .models import Salarie



def accueil_user(request):
    return render(request,"accounts/accueil.html",{})


def contact(request):
    return render(request,"accounts/contact.html",{})


def login_user(request):
    if request.method == "POST":
        username = request.POST["username"]
        password = request.POST["password"]

        user = authenticate(request, username = username, password = password)

        if user is not None :
            login(request, user)
            return redirect("accounts:accueil")
        else:
            messages.info(request, "Identifiant ou mot de passe incorrect")
    
    form = AuthenticationForm()
    return render(request, "accounts/login.html", { "form" : form})



def logout_user(request):
    logout(request)
    return redirect("accounts:accueil")


def register_user(request):
    if request.method == "POST":
        form = UserCreationForm(request.POST)

        if form.is_valid():
            form.save()
            return redirect("ressources:index")
    else:
        form = UserCreationForm()

    return render(request, "accounts/register.html", {"form":form})



# accounts/views.py
from django.contrib.auth.decorators import login_required
from .forms import EntrepriseSignUpForm, SalarieSignUpForm
from .models import Entreprise, Salarie

def entreprise_signup(request):
    if request.method == 'POST':
        form = EntrepriseSignUpForm(request.POST)
        if form.is_valid():
            user = form.save()
            login(request, user)
            return redirect('accounts:accueil')
    else:
        form = EntrepriseSignUpForm()
    return render(request, 'accounts/register.html', {'form': form})

@login_required
def add_salarie(request):
    entreprise = Entreprise.objects.get(user=request.user)
    if request.method == 'POST':
        form = SalarieSignUpForm(request.POST, entreprise=entreprise)
        if form.is_valid():
            form.save()
            return redirect('accounts:salarie_list')
    else:
        form = SalarieSignUpForm()
    return render(request, 'accounts/register_employee.html', {'form': form})

@login_required
def salarie_list(request):
    entreprise = Entreprise.objects.get(user=request.user)
    salaries = entreprise.salaries.all()

    if request.method == 'POST':
        form = SalarieSignUpForm(request.POST, entreprise=entreprise)
        if form.is_valid():
            form.save()
            return redirect('accounts:salarie_list')
    else:
        form = SalarieSignUpForm()
    return render(request, 'accounts/salarie_list.html', {'form': form,'salaries': salaries})



from django.contrib import messages


@login_required
def delete_salarie(request, salarie_id):
    try:
        book = Salarie.objects.get(pk = salarie_id)
        book.delete()
    except Salarie.DoesNotExist:
        return redirect("accounts:salarie_list")
    return redirect("accounts:salarie_list")




