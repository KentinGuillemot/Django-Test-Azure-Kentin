from django import forms
from .models import Bulletin, Bulletin_Net_Vers_Brut, Bulletin_Cout_Vers_Brut

class BulletinForm(forms.ModelForm):

    class Meta:
        model = Bulletin
        fields = '__all__'
        exclude = ['user']
        widgets = {
            'prevoyance': forms.CheckboxInput(),
            'mutuelle': forms.CheckboxInput(),
            'taux_ccss_red': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')], attrs={'required': False}),
            'admin_SAM': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')], attrs={'required': False}),
            'Exclu_ass_cho': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')], attrs={'required': False}),
            'CCPB_ouvrier': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')], attrs={'required': False}),
            'CCPB_Etam_cadre': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')], attrs={'required': False}),
            'Gens_de_Maison': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')], attrs={'required': False}),
            'type_salarie': forms.RadioSelect(choices=[('Cadre', 'Cadre'), ('Non-cadre', 'Non-cadre')]),
            'type_plafond': forms.RadioSelect(choices=[('CCSS Monaco', 'CCSS Monaco'), ('FR URSSAF', 'FR URSSAF'), ('Je ne sais pas', 'Je ne sais pas')]),
        }

    def clean(self):
        cleaned_data = super().clean()

        prevoyance = cleaned_data.get('prevoyance')
        mutuelle = cleaned_data.get('mutuelle')

        if not prevoyance:
            cleaned_data['type_plafond'] = 'Non renseigné'
            cleaned_data['taux_prev_T1_sal'] = 0
            cleaned_data['taux_prev_T1_patr'] = 0
            cleaned_data['taux_prev_T2_sal'] = 0
            cleaned_data['taux_prev_T2_patr'] = 0

        if not mutuelle:
            cleaned_data['montant_mutu_sal'] = 0
            cleaned_data['montant_mutu_patr'] = 0

        for field in ['taux_ccss_red', 'admin_SAM', 'Exclu_ass_cho', 'CCPB_ouvrier', 'CCPB_Etam_cadre', 'Gens_de_Maison']:
            if not cleaned_data.get(field):
                cleaned_data[field] = 'Non'

        return cleaned_data
    
    def clean_heures_semaine(self):
        heures_semaine = self.cleaned_data.get('heures_semaine')
        if heures_semaine > 39:
            raise forms.ValidationError("Le nombre d'heures par semaine ne peut pas dépasser 39.")
        return heures_semaine

    def clean_heures_mois(self):
        heures_mois = self.cleaned_data.get('heures_mois')
        if heures_mois > 169:
            raise forms.ValidationError("Le nombre d'heures par mois ne peut pas dépasser 169.")
        return heures_mois
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Personnaliser le message d'erreur pour tous les champs obligatoires
        for field_name, field in self.fields.items():
            if field.required:
                field.error_messages['required'] = 'Champ obligatoire.'


class BulletinForm_Net_Vers_Brut(forms.ModelForm):

    class Meta:
        model = Bulletin_Net_Vers_Brut
        fields = '__all__'
        exclude = ['user']
        widgets = {
            'prevoyance': forms.CheckboxInput(),
            'mutuelle': forms.CheckboxInput(),
            'taux_ccss_red': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')], attrs={'required': False}),
            'admin_SAM': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')], attrs={'required': False}),
            'Exclu_ass_cho': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')], attrs={'required': False}),
            'CCPB_ouvrier': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')], attrs={'required': False}),
            'CCPB_Etam_cadre': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')], attrs={'required': False}),
            'Gens_de_Maison': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')], attrs={'required': False}),
            'type_salarie': forms.RadioSelect(choices=[('Cadre', 'Cadre'), ('Non-cadre', 'Non-cadre')]),
            'type_plafond': forms.RadioSelect(choices=[('CCSS Monaco', 'CCSS Monaco'), ('FR URSSAF', 'FR URSSAF'), ('Je ne sais pas', 'Je ne sais pas')]),
            'indemnite': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')]),
        }

    def clean(self):
        cleaned_data = super().clean()

        prevoyance = cleaned_data.get('prevoyance')
        mutuelle = cleaned_data.get('mutuelle')

        if not prevoyance:
            cleaned_data['type_plafond'] = 'Non renseigné'
            cleaned_data['taux_prev_T1_sal'] = 0
            cleaned_data['taux_prev_T1_patr'] = 0
            cleaned_data['taux_prev_T2_sal'] = 0
            cleaned_data['taux_prev_T2_patr'] = 0

        if not mutuelle:
            cleaned_data['montant_mutu_sal'] = 0
            cleaned_data['montant_mutu_patr'] = 0

        for field in ['taux_ccss_red', 'admin_SAM', 'Exclu_ass_cho', 'CCPB_ouvrier', 'CCPB_Etam_cadre', 'Gens_de_Maison']:
            if not cleaned_data.get(field):
                cleaned_data[field] = 'Non'

        return cleaned_data
    
    def clean_heures_semaine(self):
        heures_semaine = self.cleaned_data.get('heures_semaine')
        if heures_semaine > 39:
            raise forms.ValidationError("Le nombre d'heures par semaine ne peut pas dépasser 39.")
        return heures_semaine

    def clean_heures_mois(self):
        heures_mois = self.cleaned_data.get('heures_mois')
        if heures_mois > 169:
            raise forms.ValidationError("Le nombre d'heures par mois ne peut pas dépasser 169.")
        return heures_mois
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Personnaliser le message d'erreur pour tous les champs obligatoires
        for field_name, field in self.fields.items():
            if field.required:
                field.error_messages['required'] = 'Champ obligatoire.'



class BulletinForm_Cout_Vers_Brut(forms.ModelForm):

    class Meta:
        model = Bulletin_Cout_Vers_Brut
        fields = '__all__'
        exclude = ['user']
        widgets = {
            'prevoyance': forms.CheckboxInput(),
            'mutuelle': forms.CheckboxInput(),
            'taux_ccss_red': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')], attrs={'required': False}),
            'admin_SAM': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')], attrs={'required': False}),
            'Exclu_ass_cho': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')], attrs={'required': False}),
            'CCPB_ouvrier': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')], attrs={'required': False}),
            'CCPB_Etam_cadre': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')], attrs={'required': False}),
            'Gens_de_Maison': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')], attrs={'required': False}),
            'type_salarie': forms.RadioSelect(choices=[('Cadre', 'Cadre'), ('Non-cadre', 'Non-cadre')]),
            'type_plafond': forms.RadioSelect(choices=[('CCSS Monaco', 'CCSS Monaco'), ('FR URSSAF', 'FR URSSAF'), ('Je ne sais pas', 'Je ne sais pas')]),
            'indemnite': forms.RadioSelect(choices=[('Oui', 'Oui'), ('Non', 'Non')]),
        }

    def clean(self):
        cleaned_data = super().clean()

        prevoyance = cleaned_data.get('prevoyance')
        mutuelle = cleaned_data.get('mutuelle')

        if not prevoyance:
            cleaned_data['type_plafond'] = 'Non renseigné'
            cleaned_data['taux_prev_T1_sal'] = 0
            cleaned_data['taux_prev_T1_patr'] = 0
            cleaned_data['taux_prev_T2_sal'] = 0
            cleaned_data['taux_prev_T2_patr'] = 0

        if not mutuelle:
            cleaned_data['montant_mutu_sal'] = 0
            cleaned_data['montant_mutu_patr'] = 0

        for field in ['taux_ccss_red', 'admin_SAM', 'Exclu_ass_cho', 'CCPB_ouvrier', 'CCPB_Etam_cadre', 'Gens_de_Maison']:
            if not cleaned_data.get(field):
                cleaned_data[field] = 'Non'

        return cleaned_data
    
    def clean_heures_semaine(self):
        heures_semaine = self.cleaned_data.get('heures_semaine')
        if heures_semaine > 39:
            raise forms.ValidationError("Le nombre d'heures par semaine ne peut pas dépasser 39.")
        return heures_semaine

    def clean_heures_mois(self):
        heures_mois = self.cleaned_data.get('heures_mois')
        if heures_mois > 169:
            raise forms.ValidationError("Le nombre d'heures par mois ne peut pas dépasser 169.")
        return heures_mois
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Personnaliser le message d'erreur pour tous les champs obligatoires
        for field_name, field in self.fields.items():
            if field.required:
                field.error_messages['required'] = 'Champ obligatoire.'
