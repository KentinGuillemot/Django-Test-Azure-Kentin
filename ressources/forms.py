from typing import Any
from django import forms
from .models import utilisateur, objet, resa, nom_famille, Condition

class ConditionForm(forms.ModelForm):
    class Meta:
        model = Condition
        fields = ['nom']

class utili_form(forms.ModelForm):
    conditions_utili_m2m = forms.ModelMultipleChoiceField(
        queryset=Condition.objects.all(),
        widget=forms.CheckboxSelectMultiple,
        required=False,
        label="Atouts"
    )

    class Meta:
        model = utilisateur
        fields = ["nom", "prenom", "email", "conditions_utili_m2m"]

class ObjetForm(forms.ModelForm):
    conditions_objet_m2m = forms.ModelMultipleChoiceField(
        queryset=Condition.objects.all(),
        widget=forms.CheckboxSelectMultiple,
        required=False,
        label="Conditions"
    )

    class Meta:
        model = objet
        fields = ['nom', 'famille1', 'famille2', 'famille3', 'famille4', 'conditions_objet_m2m']

def conditions_to_list(conditions_str):
    return [condition.strip() for condition in conditions_str.split(',')]


class ResaForm(forms.ModelForm):
    class Meta:
        model = resa
        fields = ['obj', 'user', 'date_debut', 'date_fin']
        widgets = {
            'obj': forms.HiddenInput(),
            'date_debut': forms.DateTimeInput(attrs={'type': 'datetime-local'}),
            'date_fin': forms.DateTimeInput(attrs={'type': 'datetime-local'}),
        }

    def __init__(self, *args, **kwargs):
        user_queryset = kwargs.pop('user_queryset', None)
        super().__init__(*args, **kwargs)
        if user_queryset is not None:
            self.fields['user'].queryset = user_queryset

    def clean(self):
        cleaned_data = super().clean()
        obj = cleaned_data.get('obj')
        date_debut = cleaned_data.get('date_debut')
        date_fin = cleaned_data.get('date_fin')

        if obj and date_debut and date_fin:
            new_resa = resa(obj=obj, date_debut=date_debut, date_fin=date_fin)
            conflicts = new_resa.check_conflict()
            if conflicts.exists():
                conflict_dates = "\n".join([f"{c.date_debut} to {c.date_fin}" for c in conflicts])
                raise forms.ValidationError(f"L'objet est déjà réservé pour les dates suivantes :\n{conflict_dates}")

        return cleaned_data


class Famille1Form(forms.ModelForm):
    class Meta:
        model = nom_famille
        fields = ['famille1']


class Famille2Form(forms.ModelForm):
    class Meta:
        model = nom_famille
        fields = ['famille1', 'famille2']

class Famille3Form(forms.ModelForm):
    class Meta:
        model = nom_famille
        fields = ['famille1', 'famille2', 'famille3']

class Famille4Form(forms.ModelForm):
    class Meta:
        model = nom_famille
        fields = ['famille1', 'famille2', 'famille3', 'famille4']

