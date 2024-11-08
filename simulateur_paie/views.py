from django.shortcuts import render, redirect
from .forms import BulletinForm, BulletinForm_Net_Vers_Brut, BulletinForm_Cout_Vers_Brut
import openpyxl
import os
from django.conf import settings
import xlwings as xw
from django.http import HttpResponse
from openpyxl import Workbook
from .models import Bulletin, Bulletin_Net_Vers_Brut, Net_to_Brut, Bulletin_Cout_Vers_Brut, Cout_to_Brut
from django.shortcuts import get_object_or_404  
from django.db.models import Q
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors



def accueil(request):
    return render(request, 'simulateur_paie/Accueil.html', {})

def create_bulletin(request):
    bulletin_id = 0
    if request.method == 'POST':
        form = BulletinForm(request.POST)
        if form.is_valid():

            # Enregistre le form
            bulletin_instance = form.save()
            bulletin_instance.user = request.user  # Associate with the current user
            bulletin_instance.save()

            bulletin_id = bulletin_instance.id

            # Redirect to the bulletin view
            return redirect('simulateur_paie:bulletin2', bulletin_id)
        else:
            # Affiche les erreurs du formulaire dans le terminal pour le débogage
            print(form.errors)
    else:
        form = BulletinForm()
    return render(request, 'simulateur_paie/simu_paie_form.html', {'form': form, "bulletin_ID" : bulletin_id})

def create_bulletin_Net_vers_Brut(request):
    bulletin_id = 0
    if request.method == 'POST':
        form = BulletinForm_Net_Vers_Brut(request.POST)
        if form.is_valid():

            # Enregistre le form
            bulletin_instance = form.save()
            bulletin_instance.user = request.user  # Associate with the current user
            bulletin_instance.save()
            bulletin_id = bulletin_instance.id

            # Redirect to the bulletin view
            return redirect('simulateur_paie:Net_to_Brut_calcul', bulletin_id)
    else:
        form = BulletinForm_Net_Vers_Brut()
    return render(request, 'simulateur_paie/simu_paie_form_Net_to_Brut.html', {'form': form, "bulletin_ID" : bulletin_id}) 

def create_bulletin_Cout_vers_Brut(request):
    bulletin_id = 0
    if request.method == 'POST':
        form = BulletinForm_Cout_Vers_Brut(request.POST)
        if form.is_valid():

            # Enregistre le form
            bulletin_instance = form.save()
            bulletin_instance.user = request.user  # Associate with the current user
            bulletin_instance.save()

            bulletin_id = bulletin_instance.id

            # Redirect to the bulletin view
            return redirect('simulateur_paie:Cout_to_Brut_calcul', bulletin_id)
    else:
        form = BulletinForm_Cout_Vers_Brut()
    return render(request, 'simulateur_paie/simu_paie_form_Cout_to_Brut.html', {'form': form, "bulletin_ID" : bulletin_id})




def bulletin2(request,bulletin_id):

    bulletin = Bulletin.objects.get(pk=bulletin_id)

    excel_data_V2 = [[0 for _ in range(6)] for _ in range(38)]

    # Titres des colonnes
    excel_data_V2[0][0] = "Libellé"
    excel_data_V2[0][1] = "Base"
    excel_data_V2[0][2] = "Taux salarial"
    excel_data_V2[0][3] = "Montant salarial"
    excel_data_V2[0][4] = "Taux patronal"
    excel_data_V2[0][5] = "Montant patronal"

    # Titres des lignes
    excel_data_V2[2][0] = "SALAIRE DE BASE MENSUEL"
    excel_data_V2[3][0] = "VARIABLES DE REMUNERATION"

    excel_data_V2[5][0] = "TOTAL SALAIRE BRUT"

    excel_data_V2[7][0] = "CAR"
    excel_data_V2[8][0] = "CCSS"
    excel_data_V2[9][0] = "Réduction CCSS Gens de maison (33%)"
    excel_data_V2[10][0] = "POLE EMPLOI"
    excel_data_V2[11][0] = "RETRAITE COMPLEMENTAIRE CMRC T1"
    excel_data_V2[12][0] = "RETRAITE COMPLEMENTAIRE CMRC T2"
    excel_data_V2[13][0] = "PREVOYANCE MC TRANCHE A"
    excel_data_V2[14][0] = "PREVOYANCE MC TRANCHE B"
    excel_data_V2[15][0] = "PREVOYANCE FR TRANCHE A"
    excel_data_V2[16][0] = "PREVOYANCE FR TRANCHE B"
    excel_data_V2[17][0] = "ACCIDENT DU TRAVAIL"
    excel_data_V2[18][0] = "TAXE SUR COTISATION AT"
    excel_data_V2[19][0] = "CCPB OUVRIER"
    excel_data_V2[20][0] = "CCPB ETAM/CADRE"

    excel_data_V2[22][0] = "Total cotisations"

    excel_data_V2[24][0] = "Taux de charges"

    excel_data_V2[26][0] = "Non-Soumis"
    excel_data_V2[27][0] = "Indemnité 5% Monégasque"
    excel_data_V2[28][0] = "Remboursement Transport"
    excel_data_V2[29][0] = "Tickets Restaurant"
    excel_data_V2[30][0] = "Mutuelle"
    excel_data_V2[31][0] = "Indemnité Télétravail mensuelle"
    excel_data_V2[32][0] = "Paniers"

    excel_data_V2[34][0] = "Net à payer"

    excel_data_V2[36][0] = "COUT TOTAL MENSUEL EMPLOYEUR"
    excel_data_V2[37][0] = "COUT ANNUEL TOTAL EMPLOYEUR"


    # Taux et plafonds :

    CAR = 6028 # B8
    CAR_part_salariale = 0.0685 # C8
    CAR_part_patronale = 0.0831 # E8

    CCSS = 9600 # B9
    CCSS_part_patronale = 0.1345 # E9

    CCSS_taux_reduit = CCSS # B10
    CCSS_taux_reduit_part_patronale = 0.1340 # E10

    plafond_fracais_tranche_A = 3864 # B11

    SMIC = 11.65 # B12

    Horaire_temps_plein_monegasque = 169 # B13

    plafond_plein_CMRC = 151 # E13

    pole_emploi = plafond_fracais_tranche_A * 4 # B16
    pole_emploi_part_salariale = 0.024 # C16
    pole_emploi_part_patronale = 0.0405 # E16

    CMRC_tranche1 = 3947 # B17
    CMRC_tranche1_part_salariale = 0.0401 # C17
    CMRC_tranche1_part_patronale = 0.0601 # E17

    CMRC_tranche2 = CMRC_tranche1 * 7 # B18
    CMRC_tranche2_part_salariale = 0.0972 # C18
    CMRC_tranche2_part_patronale = 0.1457 # E18

    CCPB_OUVRIER_salariale = 0.004 # C19
    CCPB_OUVRIER_patronale = 0.201 # E19

    CCPB_ETAM_CADRE_salariale = 0.0 # C20
    CCPB_ETAM_CADRE_patronale = 0.205 # E20

    prevoyance_cadre_plafond_monégasque = CCSS # B21
    prevoyance_cadre_plafond_monégasque_patronale = 0.015 # E21

    prevoyance_cadre_plafond_francais =  plafond_fracais_tranche_A # B22
    prevoyance_cadre_plafond_francais_patronale = 0.015 # E22

    taxe_cotisation_AT_patronale = 0.04 # E23

    minimu_garanti = 4.15 # B25



    #ligne 31
    excel_data_V2[2][1] = round(bulletin.heures_mois,2) # B31
    excel_data_V2[2][2] = bulletin.salaire_brut_mois / excel_data_V2[2][1] # C31
    excel_data_V2[2][3] = bulletin.salaire_brut_mois # D31

    #ligne 34
    excel_data_V2[5][3] = excel_data_V2[2][3] # D34

    #ligne 36
    if excel_data_V2[5][3] < CAR:
        excel_data_V2[7][1] = excel_data_V2[5][3] # B36
    else:
        excel_data_V2[7][1] = CAR # B36

    excel_data_V2[7][2] = CAR_part_salariale # C36
    excel_data_V2[7][3] = excel_data_V2[7][1] * excel_data_V2[7][2] # D36
    excel_data_V2[7][4] = CAR_part_patronale # E36
    excel_data_V2[7][5] = excel_data_V2[7][1] * excel_data_V2[7][4] # F36

    #ligne 37
    if bulletin.admin_SAM == "Oui":
        excel_data_V2[8][1] = CCSS # B37
    else:
        excel_data_V2[8][1] = min(excel_data_V2[5][3], CCSS) # B37

    if bulletin.taux_ccss_red == "Oui":
        excel_data_V2[8][4] = CCSS_taux_reduit_part_patronale # E37
    else:
        excel_data_V2[8][4] = CCSS_part_patronale # E37

    excel_data_V2[8][5] = excel_data_V2[8][1] * excel_data_V2[8][4] # F37

    #ligne 38
    if bulletin.Gens_de_Maison == "Oui":
        excel_data_V2[9][1] = -excel_data_V2[8][1] * 0.67 # B38
    else:
        excel_data_V2[9][1] = 0 # B38

    excel_data_V2[9][4] = CCSS_taux_reduit_part_patronale # E38
    excel_data_V2[9][5] = excel_data_V2[9][1] * excel_data_V2[9][4] # F38

    #ligne 39
    H39 = (plafond_fracais_tranche_A / 169 * excel_data_V2[2][1]) * 4 #********************************************************

    if bulletin.Exclu_ass_cho == "Oui":
        excel_data_V2[10][1] = 0 # B39
    else:
        excel_data_V2[10][1] = min(excel_data_V2[5][3], H39) # B39

    excel_data_V2[10][2] = pole_emploi_part_salariale # C39
    excel_data_V2[10][3] = excel_data_V2[10][1] * excel_data_V2[10][2] # D39
    excel_data_V2[10][4] = pole_emploi_part_patronale # E39
    excel_data_V2[10][5] = excel_data_V2[10][1] * excel_data_V2[10][4] # F39

    #ligne 40
    rounded_B31 = round(excel_data_V2[2][1])  # Arrondi de B31 à 0 décimale
    calculation = (CMRC_tranche1 / plafond_plein_CMRC) * rounded_B31 
    H40 = min(CMRC_tranche1, calculation)

    if excel_data_V2[5][3] <= H40:
        excel_data_V2[11][1] = excel_data_V2[5][3] # B40
    else:
        excel_data_V2[11][1] = H40 # B40

    excel_data_V2[11][2] = CMRC_tranche1_part_salariale # C40
    excel_data_V2[11][3] = excel_data_V2[11][1] * excel_data_V2[11][2] # D40
    excel_data_V2[11][4] = CMRC_tranche1_part_patronale # E40
    excel_data_V2[11][5] = excel_data_V2[11][1] * excel_data_V2[11][4] # F40

    #ligne 41
    H41 = H40*7

    if excel_data_V2[5][3] <= (H40 + H41):
        excel_data_V2[12][1] = excel_data_V2[5][3] - excel_data_V2[11][1] # B41
    else:
        excel_data_V2[12][1] = (H40 + H41) - excel_data_V2[11][1] # B41
    
    excel_data_V2[12][2] = CMRC_tranche2_part_salariale  # C41
    excel_data_V2[12][3] = excel_data_V2[12][1] * excel_data_V2[12][2] # D41
    excel_data_V2[12][4] = CMRC_tranche2_part_patronale # E41
    excel_data_V2[12][5] = excel_data_V2[12][1] * excel_data_V2[12][4] # F41

    #ligne 42
    B8 = "Oui" if bulletin.prevoyance else "Non"
    H42 = prevoyance_cadre_plafond_monégasque / Horaire_temps_plein_monegasque * excel_data_V2[2][1]

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[13][1] = 0 # B42
        else:
            excel_data_V2[13][1] = min(H42, excel_data_V2[5][3]) # B42
    else:
        excel_data_V2[13][1] = 0 # B42
    
    excel_data_V2[13][2] = bulletin.taux_prev_T1_sal / 100 # C42
    excel_data_V2[13][3] = excel_data_V2[13][1] * excel_data_V2[13][2]  # D42
    excel_data_V2[13][4] = bulletin.taux_prev_T1_patr / 100 # E42
    excel_data_V2[13][5] = excel_data_V2[13][1] * excel_data_V2[13][4] # F42

    #ligne 43
    H43 = H39 - H42

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[14][1] = 0 # B43
        else:
            if excel_data_V2[5][3] > H42:
                excel_data_V2[14][1] = min(excel_data_V2[5][3]-excel_data_V2[13][1],H42+H43-excel_data_V2[13][1]) # B43
            else:
                excel_data_V2[14][1] = 0 # B43
    else:
        excel_data_V2[14][1] = 0 # B43
    
    excel_data_V2[14][2] =  bulletin.taux_prev_T2_sal / 100 # C43
    excel_data_V2[14][3] =  excel_data_V2[14][1] * excel_data_V2[14][2] # D43
    excel_data_V2[14][4] =  bulletin.taux_prev_T2_patr / 100 # E43
    excel_data_V2[14][5] =  excel_data_V2[14][1] * excel_data_V2[14][4] # F43

    #ligne 44
    H46 = plafond_fracais_tranche_A / Horaire_temps_plein_monegasque * excel_data_V2[2][1]

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[15][1] = min(excel_data_V2[5][3],H46) # B44
        else:
            excel_data_V2[15][1] = 0 # B44
    else:
        excel_data_V2[15][1] = 0 # B44

    excel_data_V2[15][2] = bulletin.taux_prev_T1_sal / 100 # C44
    excel_data_V2[15][3] = excel_data_V2[15][1] * excel_data_V2[15][2]  # D44
    excel_data_V2[15][4] = bulletin.taux_prev_T1_patr / 100 # E44
    excel_data_V2[15][5] = excel_data_V2[15][1] * excel_data_V2[15][4] # F44

    #ligne 45
    H49 = H46*4

    if B8 == "Non":
        excel_data_V2[16][1] = 0 # B45
    else:
        if bulletin.type_plafond != "FR URSSAF":
            excel_data_V2[16][1]  = 0 # B45
        else:
            if excel_data_V2[5][3] > H46: 
                excel_data_V2[16][1] = min(excel_data_V2[5][3] - excel_data_V2[15][1], H49 - excel_data_V2[15][1]) # B45
            else:
                excel_data_V2[16][1] = 0 # B45

    excel_data_V2[16][2] = bulletin.taux_prev_T2_sal / 100 # C45
    excel_data_V2[16][3] = excel_data_V2[16][1] * excel_data_V2[16][2] # D45
    excel_data_V2[16][4] = bulletin.taux_prev_T2_patr / 100 # E45
    excel_data_V2[16][3] = excel_data_V2[16][1] * excel_data_V2[16][4] # F45


    #ligne 46
    excel_data_V2[17][1] =  excel_data_V2[5][3] # B46

    if bulletin.Taux_ATHT > 0:
        excel_data_V2[17][4] = bulletin.Taux_ATHT / 100 # E46
    else:
        excel_data_V2[17][4] = 1 / 100 # E46

    excel_data_V2[17][5] = excel_data_V2[17][1] * excel_data_V2[17][4] # F46

    #ligne 47
    excel_data_V2[18][1] = excel_data_V2[17][5] # B47

    excel_data_V2[18][4] = taxe_cotisation_AT_patronale # E47

    excel_data_V2[18][5] = excel_data_V2[18][1] * excel_data_V2[18][4] # F47

    #ligne 48
    if bulletin.CCPB_ouvrier == "Oui":
        excel_data_V2[19][1] = excel_data_V2[5][3] # B48
    else:
        excel_data_V2[19][1] = 0 # B48

    excel_data_V2[19][2] = CCPB_OUVRIER_salariale # C48

    excel_data_V2[19][3] = excel_data_V2[19][1] * excel_data_V2[19][2] # D48
    excel_data_V2[19][4] = CCPB_OUVRIER_patronale # E48
    excel_data_V2[19][5] = excel_data_V2[19][1] * excel_data_V2[19][4] # F48
    

    #ligne 49

    if bulletin.CCPB_Etam_cadre == "Oui":
        excel_data_V2[20][1] = excel_data_V2[5][3] # B49
    else:
        excel_data_V2[20][1] = 0 # B49

    excel_data_V2[20][2] = CCPB_ETAM_CADRE_salariale # C49 

    excel_data_V2[20][3] = excel_data_V2[20][1] * excel_data_V2[20][2] # D49
    excel_data_V2[20][4] = CCPB_ETAM_CADRE_patronale # E49
    excel_data_V2[20][5] = excel_data_V2[20][1] * excel_data_V2[20][4] # F49


    #ligne 51
    excel_data_V2[22][3] = 0
    for i in range(7,21):
        excel_data_V2[22][3] = excel_data_V2[22][3] + excel_data_V2[i][3] # D51

    excel_data_V2[22][5] = 0
    for i in range(7,21):
        excel_data_V2[22][5] = excel_data_V2[22][5] + excel_data_V2[i][5] # F51

    #ligne 53
    excel_data_V2[24][2] = excel_data_V2[22][3] / excel_data_V2[5][3] # C53

    excel_data_V2[24][4] = excel_data_V2[22][5] / excel_data_V2[5][3] # E53

    #ligne 56
    H50 = SMIC * 1.05

    if excel_data_V2[2][2] <= H50:
        excel_data_V2[27][1] = excel_data_V2[5][3] # B56
    else:
        excel_data_V2[27][1] = 0 # B56

    excel_data_V2[27][2] = 0.05 # C56

    excel_data_V2[27][3] = excel_data_V2[27][1] * excel_data_V2[27][2] # D56

    excel_data_V2[27][5] = excel_data_V2[27][3] # F56

    #ligne 57
    excel_data_V2[28][3] = bulletin.remboursement_transport # D57

    excel_data_V2[28][5] = excel_data_V2[28][3] # F57

    #ligne 58
    excel_data_V2[29][1] = bulletin.ticket_resto_prix # B58
    excel_data_V2[29][2] = -bulletin.ticket_resto_sal # C58
    excel_data_V2[29][3] = excel_data_V2[29][1] * excel_data_V2[29][2] # D58
    excel_data_V2[29][4] = bulletin.ticket_resto_patr # E58
    excel_data_V2[29][5] = excel_data_V2[29][4] * excel_data_V2[29][1] # F58

    #ligne 59
    B12 = "Oui" if bulletin.mutuelle else "Non"

    if B12 == "Oui":
        excel_data_V2[30][3] = -bulletin.montant_mutu_sal # D59
    else:
        excel_data_V2[30][3] = 0 # D59

    if B12 == "Oui":
        excel_data_V2[30][5] = bulletin.montant_mutu_patr # F59
    else:
        excel_data_V2[30][5] = 0 # F59

    #ligne 60
    excel_data_V2[31][3] = 20 # D60

    #ligne 61
    excel_data_V2[32][1] = bulletin.panier_prix # B61
    excel_data_V2[32][2] = bulletin.paniers_sal # C61
    excel_data_V2[32][3] = excel_data_V2[32][1] * excel_data_V2[32][2] # D61
    excel_data_V2[32][5] = excel_data_V2[32][3] # F61

    #ligne 62
    excel_data_V2[33][5] = 0
    for i in range(22,33):
        excel_data_V2[33][5] = excel_data_V2[33][5] + excel_data_V2[i][5] # F62

    #ligne 63
    somme = 0
    for i in range(24,33):
        somme = somme + excel_data_V2[i][3]

    excel_data_V2[34][3] = excel_data_V2[2][3] - excel_data_V2[22][3] + somme # D63

    #ligne 65
    excel_data_V2[36][1] = excel_data_V2[33][5] + excel_data_V2[2][3] # B65

    #ligne 66
    excel_data_V2[37][1] = excel_data_V2[36][1]*12 # B66



    # Formatage des datas (arrondis, %, ...)
    for i in range(7, 28):
        if excel_data_V2[i][2] is not None:
            excel_data_V2[i][2] = excel_data_V2[i][2] * 100

    for i in range(7, 25):
        if excel_data_V2[i][4] is not None:
            excel_data_V2[i][4] = excel_data_V2[i][4] * 100

    for i in range(2,38):
        for j in range(1,6):
            if (excel_data_V2[i][j] is not None) and not isinstance(excel_data_V2[i][j], str):
                if (i== 2) and (j==2):
                    excel_data_V2[i][j] = f"{excel_data_V2[i][j]:.4f}" # attention ça fait que maintenant les chiffres deviennent des strings
                else:
                    excel_data_V2[i][j] = f"{excel_data_V2[i][j]:.2f}" # attention ça fait que maintenant les chiffres deviennent des strings

                if (excel_data_V2[i][j] == f"{0.0:.2f}"):
                    excel_data_V2[i][j] = " "

                excel_data_V2[i][j] = str(excel_data_V2[i][j])

                for p in range(0, len(excel_data_V2[i][j])):
                    if (excel_data_V2[i][j][p] == ".") and (len(excel_data_V2[i][j][:p]) >= 4) and (len(excel_data_V2[i][j][:p]) < 7):
                        u = p - 3
                        excel_data_V2[i][j] = excel_data_V2[i][j][:u] + " " + excel_data_V2[i][j][u:]
                                

    excel_data_V2[34][3] = excel_data_V2[34][3] + " €"
    excel_data_V2[36][1] = excel_data_V2[36][1] + " €"
    excel_data_V2[37][1] = excel_data_V2[37][1] + " €"

    for i in range(7, 28):
        if excel_data_V2[i][2] is not None and excel_data_V2[i][2] != " ":
            excel_data_V2[i][2] = str(excel_data_V2[i][2]) + "%"

    for i in range(7, 25):
        if excel_data_V2[i][4] is not None and excel_data_V2[i][4] != " ":
            excel_data_V2[i][4] = str(excel_data_V2[i][4]) + "%"


    #Suppression des lignes entierement à 0
    # def remove_zero_rows_except_first_column(matrix):
    #     return [row for row in matrix if any(cell != f" " for cell in row[1:])]

    # excel_data_V2 = remove_zero_rows_except_first_column(excel_data_V2)
    # Pour l'instant on laisse en com parce que faut adapter le CSS pour que ça rende bien vu que la c'est sur des chiffres en dur


    context = {
        'excel_data': excel_data_V2,
        'bulletin' : bulletin,
    }

    return render(request, 'simulateur_paie/bulletin.html', context)


def simu_net_avec_brut_estime(brut_estime,bulletin_id):

    bulletin = Bulletin_Net_Vers_Brut.objects.get(pk=bulletin_id)

    excel_data_V2 = [[0 for _ in range(6)] for _ in range(38)]

    # Titres des colonnes
    excel_data_V2[0][0] = "Libellé"
    excel_data_V2[0][1] = "Base"
    excel_data_V2[0][2] = "Taux salarial"
    excel_data_V2[0][3] = "Montant salarial"
    excel_data_V2[0][4] = "Taux patronal"
    excel_data_V2[0][5] = "Montant patronal"

    # Titres des lignes
    excel_data_V2[2][0] = "SALAIRE DE BASE MENSUEL"
    excel_data_V2[3][0] = "VARIABLES DE REMUNERATION"

    excel_data_V2[5][0] = "TOTAL SALAIRE BRUT"

    excel_data_V2[7][0] = "CAR"
    excel_data_V2[8][0] = "CCSS"
    excel_data_V2[9][0] = "Réduction CCSS Gens de maison (33%)"
    excel_data_V2[10][0] = "POLE EMPLOI"
    excel_data_V2[11][0] = "RETRAITE COMPLEMENTAIRE CMRC T1"
    excel_data_V2[12][0] = "RETRAITE COMPLEMENTAIRE CMRC T2"
    excel_data_V2[13][0] = "PREVOYANCE MC TRANCHE A"
    excel_data_V2[14][0] = "PREVOYANCE MC TRANCHE B"
    excel_data_V2[15][0] = "PREVOYANCE FR TRANCHE A"
    excel_data_V2[16][0] = "PREVOYANCE FR TRANCHE B"
    excel_data_V2[17][0] = "ACCIDENT DU TRAVAIL"
    excel_data_V2[18][0] = "TAXE SUR COTISATION AT"
    excel_data_V2[19][0] = "CCPB OUVRIER"
    excel_data_V2[20][0] = "CCPB ETAM/CADRE"

    excel_data_V2[22][0] = "Total cotisations"

    excel_data_V2[24][0] = "Taux de charges"

    excel_data_V2[26][0] = "Non-Soumis"
    excel_data_V2[27][0] = "Indemnité 5% Monégasque"
    excel_data_V2[28][0] = "Remboursement Transport"
    excel_data_V2[29][0] = "Tickets Restaurant"
    excel_data_V2[30][0] = "Mutuelle"
    excel_data_V2[31][0] = "Indemnité Télétravail mensuelle"
    excel_data_V2[32][0] = "Paniers"

    excel_data_V2[34][0] = "Net à payer"

    excel_data_V2[36][0] = "COUT TOTAL MENSUEL EMPLOYEUR"
    excel_data_V2[37][0] = "COUT ANNUEL TOTAL EMPLOYEUR"


    # Taux et plafonds :

    CAR = 6028 # B8
    CAR_part_salariale = 0.0685 # C8
    CAR_part_patronale = 0.0831 # E8

    CCSS = 9600 # B9
    CCSS_part_patronale = 0.1345 # E9

    CCSS_taux_reduit = CCSS # B10
    CCSS_taux_reduit_part_patronale = 0.1340 # E10

    plafond_fracais_tranche_A = 3864 # B11

    SMIC = 11.65 # B12

    Horaire_temps_plein_monegasque = 169 # B13

    plafond_plein_CMRC = 151 # E13

    pole_emploi = plafond_fracais_tranche_A * 4 # B16
    pole_emploi_part_salariale = 0.024 # C16
    pole_emploi_part_patronale = 0.0405 # E16

    CMRC_tranche1 = 3947 # B17
    CMRC_tranche1_part_salariale = 0.0401 # C17
    CMRC_tranche1_part_patronale = 0.0601 # E17

    CMRC_tranche2 = CMRC_tranche1 * 7 # B18
    CMRC_tranche2_part_salariale = 0.0972 # C18
    CMRC_tranche2_part_patronale = 0.1457 # E18

    CCPB_OUVRIER_salariale = 0.004 # C19
    CCPB_OUVRIER_patronale = 0.201 # E19

    CCPB_ETAM_CADRE_salariale = 0.0 # C20
    CCPB_ETAM_CADRE_patronale = 0.205 # E20

    prevoyance_cadre_plafond_monégasque = CCSS # B21
    prevoyance_cadre_plafond_monégasque_patronale = 0.015 # E21

    prevoyance_cadre_plafond_francais =  plafond_fracais_tranche_A # B22
    prevoyance_cadre_plafond_francais_patronale = 0.015 # E22

    taxe_cotisation_AT_patronale = 0.04 # E23

    minimu_garanti = 4.15 # B25



    #ligne 31
    excel_data_V2[2][1] = round(bulletin.heures_mois,2) # B31
    excel_data_V2[2][2] = brut_estime / excel_data_V2[2][1] # C31
    excel_data_V2[2][3] = brut_estime # D31

    #ligne 34
    excel_data_V2[5][3] = excel_data_V2[2][3] # D34

    #ligne 36
    if excel_data_V2[5][3] < CAR:
        excel_data_V2[7][1] = excel_data_V2[5][3] # B36
    else:
        excel_data_V2[7][1] = CAR # B36

    excel_data_V2[7][2] = CAR_part_salariale # C36
    excel_data_V2[7][3] = excel_data_V2[7][1] * excel_data_V2[7][2] # D36
    excel_data_V2[7][4] = CAR_part_patronale # E36
    excel_data_V2[7][5] = excel_data_V2[7][1] * excel_data_V2[7][4] # F36

    #ligne 37
    if bulletin.admin_SAM == "Oui":
        excel_data_V2[8][1] = CCSS # B37
    else:
        excel_data_V2[8][1] = min(excel_data_V2[5][3], CCSS) # B37

    if bulletin.taux_ccss_red == "Oui":
        excel_data_V2[8][4] = CCSS_taux_reduit_part_patronale # E37
    else:
        excel_data_V2[8][4] = CCSS_part_patronale # E37

    excel_data_V2[8][5] = excel_data_V2[8][1] * excel_data_V2[8][4] # F37

    #ligne 38
    if bulletin.Gens_de_Maison == "Oui":
        excel_data_V2[9][1] = -excel_data_V2[8][1] * 0.67 # B38
    else:
        excel_data_V2[9][1] = 0 # B38

    excel_data_V2[9][4] = CCSS_taux_reduit_part_patronale # E38
    excel_data_V2[9][5] = excel_data_V2[9][1] * excel_data_V2[9][4] # F38

    #ligne 39
    H39 = (plafond_fracais_tranche_A / 169 * excel_data_V2[2][1]) * 4 #********************************************************

    if bulletin.Exclu_ass_cho == "Oui":
        excel_data_V2[10][1] = 0 # B39
    else:
        excel_data_V2[10][1] = min(excel_data_V2[5][3], H39) # B39

    excel_data_V2[10][2] = pole_emploi_part_salariale # C39
    excel_data_V2[10][3] = excel_data_V2[10][1] * excel_data_V2[10][2] # D39
    excel_data_V2[10][4] = pole_emploi_part_patronale # E39
    excel_data_V2[10][5] = excel_data_V2[10][1] * excel_data_V2[10][4] # F39

    #ligne 40
    rounded_B31 = round(excel_data_V2[2][1])  # Arrondi de B31 à 0 décimale
    calculation = (CMRC_tranche1 / plafond_plein_CMRC) * rounded_B31 
    H40 = min(CMRC_tranche1, calculation)

    if excel_data_V2[5][3] <= H40:
        excel_data_V2[11][1] = excel_data_V2[5][3] # B40
    else:
        excel_data_V2[11][1] = H40 # B40

    excel_data_V2[11][2] = CMRC_tranche1_part_salariale # C40
    excel_data_V2[11][3] = excel_data_V2[11][1] * excel_data_V2[11][2] # D40
    excel_data_V2[11][4] = CMRC_tranche1_part_patronale # E40
    excel_data_V2[11][5] = excel_data_V2[11][1] * excel_data_V2[11][4] # F40

    #ligne 41
    H41 = H40*7

    if excel_data_V2[5][3] <= (H40 + H41):
        excel_data_V2[12][1] = excel_data_V2[5][3] - excel_data_V2[11][1] # B41
    else:
        excel_data_V2[12][1] = (H40 + H41) - excel_data_V2[11][1] # B41
    
    excel_data_V2[12][2] = CMRC_tranche2_part_salariale  # C41
    excel_data_V2[12][3] = excel_data_V2[12][1] * excel_data_V2[12][2] # D41
    excel_data_V2[12][4] = CMRC_tranche2_part_patronale # E41
    excel_data_V2[12][5] = excel_data_V2[12][1] * excel_data_V2[12][4] # F41

    #ligne 42
    B8 = "Oui" if bulletin.prevoyance else "Non"
    H42 = prevoyance_cadre_plafond_monégasque / Horaire_temps_plein_monegasque * excel_data_V2[2][1]

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[13][1] = 0 # B42
        else:
            excel_data_V2[13][1] = min(H42, excel_data_V2[5][3]) # B42
    else:
        excel_data_V2[13][1] = 0 # B42
    
    excel_data_V2[13][2] = bulletin.taux_prev_T1_sal / 100 # C42
    excel_data_V2[13][3] = excel_data_V2[13][1] * excel_data_V2[13][2]  # D42
    excel_data_V2[13][4] = bulletin.taux_prev_T1_patr / 100 # E42
    excel_data_V2[13][5] = excel_data_V2[13][1] * excel_data_V2[13][4] # F42

    #ligne 43
    H43 = H39 - H42

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[14][1] = 0 # B43
        else:
            if excel_data_V2[5][3] > H42:
                excel_data_V2[14][1] = min(excel_data_V2[5][3]-excel_data_V2[13][1],H42+H43-excel_data_V2[13][1]) # B43
            else:
                excel_data_V2[14][1] = 0 # B43
    else:
        excel_data_V2[14][1] = 0 # B43
    
    excel_data_V2[14][2] =  bulletin.taux_prev_T2_sal / 100 # C43
    excel_data_V2[14][3] =  excel_data_V2[14][1] * excel_data_V2[14][2] # D43
    excel_data_V2[14][4] =  bulletin.taux_prev_T2_patr / 100 # E43
    excel_data_V2[14][5] =  excel_data_V2[14][1] * excel_data_V2[14][4] # F43

    #ligne 44
    H46 = plafond_fracais_tranche_A / Horaire_temps_plein_monegasque * excel_data_V2[2][1]

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[15][1] = min(excel_data_V2[5][3],H46) # B44
        else:
            excel_data_V2[15][1] = 0 # B44
    else:
        excel_data_V2[15][1] = 0 # B44

    excel_data_V2[15][2] = bulletin.taux_prev_T1_sal / 100 # C44
    excel_data_V2[15][3] = excel_data_V2[15][1] * excel_data_V2[15][2]  # D44
    excel_data_V2[15][4] = bulletin.taux_prev_T1_patr / 100 # E44
    excel_data_V2[15][5] = excel_data_V2[15][1] * excel_data_V2[15][4] # F44

    #ligne 45
    H49 = H46*4

    if B8 == "Non":
        excel_data_V2[16][1] = 0 # B45
    else:
        if bulletin.type_plafond != "FR URSSAF":
            excel_data_V2[16][1]  = 0 # B45
        else:
            if excel_data_V2[5][3] > H46: 
                excel_data_V2[16][1] = min(excel_data_V2[5][3] - excel_data_V2[15][1], H49 - excel_data_V2[15][1]) # B45
            else:
                excel_data_V2[16][1] = 0 # B45

    excel_data_V2[16][2] = bulletin.taux_prev_T2_sal / 100 # C45
    excel_data_V2[16][3] = excel_data_V2[16][1] * excel_data_V2[16][2] # D45
    excel_data_V2[16][4] = bulletin.taux_prev_T2_patr / 100 # E45
    excel_data_V2[16][3] = excel_data_V2[16][1] * excel_data_V2[16][4] # F45


    #ligne 46
    excel_data_V2[17][1] =  excel_data_V2[5][3] # B46

    if bulletin.Taux_ATHT > 0:
        excel_data_V2[17][4] = bulletin.Taux_ATHT / 100 # E46
    else:
        excel_data_V2[17][4] = 1 / 100 # E46

    excel_data_V2[17][5] = excel_data_V2[17][1] * excel_data_V2[17][4] # F46

    #ligne 47
    excel_data_V2[18][1] = excel_data_V2[17][5] # B47

    excel_data_V2[18][4] = taxe_cotisation_AT_patronale # E47

    excel_data_V2[18][5] = excel_data_V2[18][1] * excel_data_V2[18][4] # F47

    #ligne 48
    if bulletin.CCPB_ouvrier == "Oui":
        excel_data_V2[19][1] = excel_data_V2[5][3] # B48
    else:
        excel_data_V2[19][1] = 0 # B48

    excel_data_V2[19][2] = CCPB_OUVRIER_salariale # C48

    excel_data_V2[19][3] = excel_data_V2[19][1] * excel_data_V2[19][2] # D48
    excel_data_V2[19][4] = CCPB_OUVRIER_patronale # E48
    excel_data_V2[19][5] = excel_data_V2[19][1] * excel_data_V2[19][4] # F48
    

    #ligne 49

    if bulletin.CCPB_Etam_cadre == "Oui":
        excel_data_V2[20][1] = excel_data_V2[5][3] # B49
    else:
        excel_data_V2[20][1] = 0 # B49

    excel_data_V2[20][2] = CCPB_ETAM_CADRE_salariale # C49 

    excel_data_V2[20][3] = excel_data_V2[20][1] * excel_data_V2[20][2] # D49
    excel_data_V2[20][4] = CCPB_ETAM_CADRE_patronale # E49
    excel_data_V2[20][5] = excel_data_V2[20][1] * excel_data_V2[20][4] # F49


    #ligne 51
    excel_data_V2[22][3] = 0
    for i in range(7,21):
        excel_data_V2[22][3] = excel_data_V2[22][3] + excel_data_V2[i][3] # D51

    excel_data_V2[22][5] = 0
    for i in range(7,21):
        excel_data_V2[22][5] = excel_data_V2[22][5] + excel_data_V2[i][5] # F51

    #ligne 53
    excel_data_V2[24][2] = excel_data_V2[22][3] / excel_data_V2[5][3] # C53

    excel_data_V2[24][4] = excel_data_V2[22][5] / excel_data_V2[5][3] # E53

    #ligne 56
    H50 = SMIC * 1.05

    
    if excel_data_V2[2][2] <= H50:
        excel_data_V2[27][1] = excel_data_V2[5][3] # B56
    else:
        excel_data_V2[27][1] = 0 # B56
    

    excel_data_V2[27][2] = 0.05 # C56

    excel_data_V2[27][3] = excel_data_V2[27][1] * excel_data_V2[27][2] # D56

    excel_data_V2[27][5] = excel_data_V2[27][3] # F56

    #ligne 57
    excel_data_V2[28][3] = bulletin.remboursement_transport # D57

    excel_data_V2[28][5] = excel_data_V2[28][3] # F57

    #ligne 58
    excel_data_V2[29][1] = bulletin.ticket_resto_prix # B58
    excel_data_V2[29][2] = -bulletin.ticket_resto_sal # C58
    excel_data_V2[29][3] = excel_data_V2[29][1] * excel_data_V2[29][2] # D58
    excel_data_V2[29][4] = bulletin.ticket_resto_patr # E58
    excel_data_V2[29][5] = excel_data_V2[29][4] * excel_data_V2[29][1] # F58

    #ligne 59
    B12 = "Oui" if bulletin.mutuelle else "Non"

    if B12 == "Oui":
        excel_data_V2[30][3] = -bulletin.montant_mutu_sal # D59
    else:
        excel_data_V2[30][3] = 0 # D59

    if B12 == "Oui":
        excel_data_V2[30][5] = bulletin.montant_mutu_patr # F59
    else:
        excel_data_V2[30][5] = 0 # F59

    #ligne 60
    excel_data_V2[31][3] = 20 # D60

    #ligne 61
    excel_data_V2[32][1] = bulletin.panier_prix # B61
    excel_data_V2[32][2] = bulletin.paniers_sal # C61
    excel_data_V2[32][3] = excel_data_V2[32][1] * excel_data_V2[32][2] # D61
    excel_data_V2[32][5] = excel_data_V2[32][3] # F61

    #ligne 62
    excel_data_V2[33][5] = 0
    for i in range(22,33):
        excel_data_V2[33][5] = excel_data_V2[33][5] + excel_data_V2[i][5] # F62

    #ligne 63
    somme = 0
    for i in range(24,33):
        somme = somme + excel_data_V2[i][3]

    excel_data_V2[34][3] = excel_data_V2[2][3] - excel_data_V2[22][3] + somme # D63


    excel_data_V2[34][3] = round(excel_data_V2[34][3],2)

    return excel_data_V2[34][3]



def calculer_net_vers_brut(request, bulletin_id):

    bulletin_res = Bulletin_Net_Vers_Brut.objects.get(pk=bulletin_id)
    salaire_net_souhaité = bulletin_res.salaire_net_mois
    indemnite = bulletin_res.indemnite
    heures = bulletin_res.heures_mois
    H50 = 11.65 * 1.05

    # Hypothèse initiale : net * 1.13 pour estimer un brut de départ
    brut_estime = salaire_net_souhaité * 1.13
    taux_horaire = brut_estime/heures
    
    if indemnite == "Oui":
        i = 0.01
        while taux_horaire > H50:
            brut_estime =  salaire_net_souhaité * (1.13-i)
            taux_horaire = brut_estime/heures
            i = i + 0.01

    
    max_iter=10000000  # 10 000 000

    for i in range(max_iter):
       
        salaire_net_simulé = simu_net_avec_brut_estime(brut_estime,bulletin_id)
        
        # Comparer le net simulé au net souhaité
        différence = salaire_net_souhaité - salaire_net_simulé
        
        # Si la différence est inférieure à la tolérance, on a trouvé une solution
        if abs(différence) == 0:
            break
        
        # Ajuster le brut estimé en fonction de la différence
        brut_estime += différence * 0.1  # Ajustement itératif léger

    brut_estime = round(brut_estime,1)

    Net_to_Brut.objects.create(salaire_brut=brut_estime,salaire_net_mois=salaire_net_souhaité, bulletin=bulletin_res)


    excel_data_V2 = [[0 for _ in range(6)] for _ in range(38)]

    # Titres des colonnes
    excel_data_V2[0][0] = "Libellé"
    excel_data_V2[0][1] = "Base"
    excel_data_V2[0][2] = "Taux salarial"
    excel_data_V2[0][3] = "Montant salarial"
    excel_data_V2[0][4] = "Taux patronal"
    excel_data_V2[0][5] = "Montant patronal"

    # Titres des lignes
    excel_data_V2[2][0] = "SALAIRE DE BASE MENSUEL"
    excel_data_V2[3][0] = "VARIABLES DE REMUNERATION"

    excel_data_V2[5][0] = "TOTAL SALAIRE BRUT"

    excel_data_V2[7][0] = "CAR"
    excel_data_V2[8][0] = "CCSS"
    excel_data_V2[9][0] = "Réduction CCSS Gens de maison (33%)"
    excel_data_V2[10][0] = "POLE EMPLOI"
    excel_data_V2[11][0] = "RETRAITE COMPLEMENTAIRE CMRC T1"
    excel_data_V2[12][0] = "RETRAITE COMPLEMENTAIRE CMRC T2"
    excel_data_V2[13][0] = "PREVOYANCE MC TRANCHE A"
    excel_data_V2[14][0] = "PREVOYANCE MC TRANCHE B"
    excel_data_V2[15][0] = "PREVOYANCE FR TRANCHE A"
    excel_data_V2[16][0] = "PREVOYANCE FR TRANCHE B"
    excel_data_V2[17][0] = "ACCIDENT DU TRAVAIL"
    excel_data_V2[18][0] = "TAXE SUR COTISATION AT"
    excel_data_V2[19][0] = "CCPB OUVRIER"
    excel_data_V2[20][0] = "CCPB ETAM/CADRE"

    excel_data_V2[22][0] = "Total cotisations"

    excel_data_V2[24][0] = "Taux de charges"

    excel_data_V2[26][0] = "Non-Soumis"
    excel_data_V2[27][0] = "Indemnité 5% Monégasque"
    excel_data_V2[28][0] = "Remboursement Transport"
    excel_data_V2[29][0] = "Tickets Restaurant"
    excel_data_V2[30][0] = "Mutuelle"
    excel_data_V2[31][0] = "Indemnité Télétravail mensuelle"
    excel_data_V2[32][0] = "Paniers"

    excel_data_V2[34][0] = "Net à payer"

    excel_data_V2[36][0] = "COUT TOTAL MENSUEL EMPLOYEUR"
    excel_data_V2[37][0] = "COUT ANNUEL TOTAL EMPLOYEUR"


    # Taux et plafonds :

    CAR = 6028 # B8
    CAR_part_salariale = 0.0685 # C8
    CAR_part_patronale = 0.0831 # E8

    CCSS = 9600 # B9
    CCSS_part_patronale = 0.1345 # E9

    CCSS_taux_reduit = CCSS # B10
    CCSS_taux_reduit_part_patronale = 0.1340 # E10

    plafond_fracais_tranche_A = 3864 # B11

    SMIC = 11.65 # B12

    Horaire_temps_plein_monegasque = 169 # B13

    plafond_plein_CMRC = 151 # E13

    pole_emploi = plafond_fracais_tranche_A * 4 # B16
    pole_emploi_part_salariale = 0.024 # C16
    pole_emploi_part_patronale = 0.0405 # E16

    CMRC_tranche1 = 3947 # B17
    CMRC_tranche1_part_salariale = 0.0401 # C17
    CMRC_tranche1_part_patronale = 0.0601 # E17

    CMRC_tranche2 = CMRC_tranche1 * 7 # B18
    CMRC_tranche2_part_salariale = 0.0972 # C18
    CMRC_tranche2_part_patronale = 0.1457 # E18

    CCPB_OUVRIER_salariale = 0.004 # C19
    CCPB_OUVRIER_patronale = 0.201 # E19

    CCPB_ETAM_CADRE_salariale = 0.0 # C20
    CCPB_ETAM_CADRE_patronale = 0.205 # E20

    prevoyance_cadre_plafond_monégasque = CCSS # B21
    prevoyance_cadre_plafond_monégasque_patronale = 0.015 # E21

    prevoyance_cadre_plafond_francais =  plafond_fracais_tranche_A # B22
    prevoyance_cadre_plafond_francais_patronale = 0.015 # E22

    taxe_cotisation_AT_patronale = 0.04 # E23

    minimu_garanti = 4.15 # B25



    #ligne 31
    excel_data_V2[2][1] = round(bulletin_res.heures_mois,2) # B31
    excel_data_V2[2][2] = brut_estime / excel_data_V2[2][1] # C31
    excel_data_V2[2][3] = brut_estime # D31

    #ligne 34
    excel_data_V2[5][3] = excel_data_V2[2][3] # D34

    #ligne 36
    if excel_data_V2[5][3] < CAR:
        excel_data_V2[7][1] = excel_data_V2[5][3] # B36
    else:
        excel_data_V2[7][1] = CAR # B36

    excel_data_V2[7][2] = CAR_part_salariale # C36
    excel_data_V2[7][3] = excel_data_V2[7][1] * excel_data_V2[7][2] # D36
    excel_data_V2[7][4] = CAR_part_patronale # E36
    excel_data_V2[7][5] = excel_data_V2[7][1] * excel_data_V2[7][4] # F36

    #ligne 37
    if bulletin_res.admin_SAM == "Oui":
        excel_data_V2[8][1] = CCSS # B37
    else:
        excel_data_V2[8][1] = min(excel_data_V2[5][3], CCSS) # B37

    if bulletin_res.taux_ccss_red == "Oui":
        excel_data_V2[8][4] = CCSS_taux_reduit_part_patronale # E37
    else:
        excel_data_V2[8][4] = CCSS_part_patronale # E37

    excel_data_V2[8][5] = excel_data_V2[8][1] * excel_data_V2[8][4] # F37

    #ligne 38
    if bulletin_res.Gens_de_Maison == "Oui":
        excel_data_V2[9][1] = -excel_data_V2[8][1] * 0.67 # B38
    else:
        excel_data_V2[9][1] = 0 # B38

    excel_data_V2[9][4] = CCSS_taux_reduit_part_patronale # E38
    excel_data_V2[9][5] = excel_data_V2[9][1] * excel_data_V2[9][4] # F38

    #ligne 39
    H39 = (plafond_fracais_tranche_A / 169 * excel_data_V2[2][1]) * 4 #********************************************************

    if bulletin_res.Exclu_ass_cho == "Oui":
        excel_data_V2[10][1] = 0 # B39
    else:
        excel_data_V2[10][1] = min(excel_data_V2[5][3], H39) # B39

    excel_data_V2[10][2] = pole_emploi_part_salariale # C39
    excel_data_V2[10][3] = excel_data_V2[10][1] * excel_data_V2[10][2] # D39
    excel_data_V2[10][4] = pole_emploi_part_patronale # E39
    excel_data_V2[10][5] = excel_data_V2[10][1] * excel_data_V2[10][4] # F39

    #ligne 40
    rounded_B31 = round(excel_data_V2[2][1])  # Arrondi de B31 à 0 décimale
    calculation = (CMRC_tranche1 / plafond_plein_CMRC) * rounded_B31 
    H40 = min(CMRC_tranche1, calculation)

    if excel_data_V2[5][3] <= H40:
        excel_data_V2[11][1] = excel_data_V2[5][3] # B40
    else:
        excel_data_V2[11][1] = H40 # B40

    excel_data_V2[11][2] = CMRC_tranche1_part_salariale # C40
    excel_data_V2[11][3] = excel_data_V2[11][1] * excel_data_V2[11][2] # D40
    excel_data_V2[11][4] = CMRC_tranche1_part_patronale # E40
    excel_data_V2[11][5] = excel_data_V2[11][1] * excel_data_V2[11][4] # F40

    #ligne 41
    H41 = H40*7

    if excel_data_V2[5][3] <= (H40 + H41):
        excel_data_V2[12][1] = excel_data_V2[5][3] - excel_data_V2[11][1] # B41
    else:
        excel_data_V2[12][1] = (H40 + H41) - excel_data_V2[11][1] # B41
    
    excel_data_V2[12][2] = CMRC_tranche2_part_salariale  # C41
    excel_data_V2[12][3] = excel_data_V2[12][1] * excel_data_V2[12][2] # D41
    excel_data_V2[12][4] = CMRC_tranche2_part_patronale # E41
    excel_data_V2[12][5] = excel_data_V2[12][1] * excel_data_V2[12][4] # F41

    #ligne 42
    B8 = "Oui" if bulletin_res.prevoyance else "Non"
    H42 = prevoyance_cadre_plafond_monégasque / Horaire_temps_plein_monegasque * excel_data_V2[2][1]

    if B8 == "Oui":
        if bulletin_res.type_plafond == "FR URSSAF":
            excel_data_V2[13][1] = 0 # B42
        else:
            excel_data_V2[13][1] = min(H42, excel_data_V2[5][3]) # B42
    else:
        excel_data_V2[13][1] = 0 # B42
    
    excel_data_V2[13][2] = bulletin_res.taux_prev_T1_sal / 100 # C42
    excel_data_V2[13][3] = excel_data_V2[13][1] * excel_data_V2[13][2]  # D42
    excel_data_V2[13][4] = bulletin_res.taux_prev_T1_patr / 100 # E42
    excel_data_V2[13][5] = excel_data_V2[13][1] * excel_data_V2[13][4] # F42

    #ligne 43
    H43 = H39 - H42

    if B8 == "Oui":
        if bulletin_res.type_plafond == "FR URSSAF":
            excel_data_V2[14][1] = 0 # B43
        else:
            if excel_data_V2[5][3] > H42:
                excel_data_V2[14][1] = min(excel_data_V2[5][3]-excel_data_V2[13][1],H42+H43-excel_data_V2[13][1]) # B43
            else:
                excel_data_V2[14][1] = 0 # B43
    else:
        excel_data_V2[14][1] = 0 # B43
    
    excel_data_V2[14][2] =  bulletin_res.taux_prev_T2_sal / 100 # C43
    excel_data_V2[14][3] =  excel_data_V2[14][1] * excel_data_V2[14][2] # D43
    excel_data_V2[14][4] =  bulletin_res.taux_prev_T2_patr / 100 # E43
    excel_data_V2[14][5] =  excel_data_V2[14][1] * excel_data_V2[14][4] # F43

    #ligne 44
    H46 = plafond_fracais_tranche_A / Horaire_temps_plein_monegasque * excel_data_V2[2][1]

    if B8 == "Oui":
        if bulletin_res.type_plafond == "FR URSSAF":
            excel_data_V2[15][1] = min(excel_data_V2[5][3],H46) # B44
        else:
            excel_data_V2[15][1] = 0 # B44
    else:
        excel_data_V2[15][1] = 0 # B44

    excel_data_V2[15][2] = bulletin_res.taux_prev_T1_sal / 100 # C44
    excel_data_V2[15][3] = excel_data_V2[15][1] * excel_data_V2[15][2]  # D44
    excel_data_V2[15][4] = bulletin_res.taux_prev_T1_patr / 100 # E44
    excel_data_V2[15][5] = excel_data_V2[15][1] * excel_data_V2[15][4] # F44

    #ligne 45
    H49 = H46*4

    if B8 == "Non":
        excel_data_V2[16][1] = 0 # B45
    else:
        if bulletin_res.type_plafond != "FR URSSAF":
            excel_data_V2[16][1]  = 0 # B45
        else:
            if excel_data_V2[5][3] > H46: 
                excel_data_V2[16][1] = min(excel_data_V2[5][3] - excel_data_V2[15][1], H49 - excel_data_V2[15][1]) # B45
            else:
                excel_data_V2[16][1] = 0 # B45

    excel_data_V2[16][2] = bulletin_res.taux_prev_T2_sal / 100 # C45
    excel_data_V2[16][3] = excel_data_V2[16][1] * excel_data_V2[16][2] # D45
    excel_data_V2[16][4] = bulletin_res.taux_prev_T2_patr / 100 # E45
    excel_data_V2[16][3] = excel_data_V2[16][1] * excel_data_V2[16][4] # F45


    #ligne 46
    excel_data_V2[17][1] =  excel_data_V2[5][3] # B46

    if bulletin_res.Taux_ATHT > 0:
        excel_data_V2[17][4] = bulletin_res.Taux_ATHT / 100 # E46
    else:
        excel_data_V2[17][4] = 1 / 100 # E46

    excel_data_V2[17][5] = excel_data_V2[17][1] * excel_data_V2[17][4] # F46

    #ligne 47
    excel_data_V2[18][1] = excel_data_V2[17][5] # B47

    excel_data_V2[18][4] = taxe_cotisation_AT_patronale # E47

    excel_data_V2[18][5] = excel_data_V2[18][1] * excel_data_V2[18][4] # F47

    #ligne 48
    if bulletin_res.CCPB_ouvrier == "Oui":
        excel_data_V2[19][1] = excel_data_V2[5][3] # B48
    else:
        excel_data_V2[19][1] = 0 # B48

    excel_data_V2[19][2] = CCPB_OUVRIER_salariale # C48

    excel_data_V2[19][3] = excel_data_V2[19][1] * excel_data_V2[19][2] # D48
    excel_data_V2[19][4] = CCPB_OUVRIER_patronale # E48
    excel_data_V2[19][5] = excel_data_V2[19][1] * excel_data_V2[19][4] # F48
    

    #ligne 49

    if bulletin_res.CCPB_Etam_cadre == "Oui":
        excel_data_V2[20][1] = excel_data_V2[5][3] # B49
    else:
        excel_data_V2[20][1] = 0 # B49

    excel_data_V2[20][2] = CCPB_ETAM_CADRE_salariale # C49 

    excel_data_V2[20][3] = excel_data_V2[20][1] * excel_data_V2[20][2] # D49
    excel_data_V2[20][4] = CCPB_ETAM_CADRE_patronale # E49
    excel_data_V2[20][5] = excel_data_V2[20][1] * excel_data_V2[20][4] # F49


    #ligne 51
    excel_data_V2[22][3] = 0
    for i in range(7,21):
        excel_data_V2[22][3] = excel_data_V2[22][3] + excel_data_V2[i][3] # D51

    excel_data_V2[22][5] = 0
    for i in range(7,21):
        excel_data_V2[22][5] = excel_data_V2[22][5] + excel_data_V2[i][5] # F51

    #ligne 53
    excel_data_V2[24][2] = excel_data_V2[22][3] / excel_data_V2[5][3] # C53

    excel_data_V2[24][4] = excel_data_V2[22][5] / excel_data_V2[5][3] # E53

    #ligne 56
    H50 = SMIC * 1.05

    if excel_data_V2[2][2] <= H50:
        excel_data_V2[27][1] = excel_data_V2[5][3] # B56
    else:
        excel_data_V2[27][1] = 0 # B56

    excel_data_V2[27][2] = 0.05 # C56

    excel_data_V2[27][3] = excel_data_V2[27][1] * excel_data_V2[27][2] # D56

    excel_data_V2[27][5] = excel_data_V2[27][3] # F56

    #ligne 57
    excel_data_V2[28][3] = bulletin_res.remboursement_transport # D57

    excel_data_V2[28][5] = excel_data_V2[28][3] # F57

    #ligne 58
    excel_data_V2[29][1] = bulletin_res.ticket_resto_prix # B58
    excel_data_V2[29][2] = -bulletin_res.ticket_resto_sal # C58
    excel_data_V2[29][3] = excel_data_V2[29][1] * excel_data_V2[29][2] # D58
    excel_data_V2[29][4] = bulletin_res.ticket_resto_patr # E58
    excel_data_V2[29][5] = excel_data_V2[29][4] * excel_data_V2[29][1] # F58

    #ligne 59
    B12 = "Oui" if bulletin_res.mutuelle else "Non"

    if B12 == "Oui":
        excel_data_V2[30][3] = -bulletin_res.montant_mutu_sal # D59
    else:
        excel_data_V2[30][3] = 0 # D59

    if B12 == "Oui":
        excel_data_V2[30][5] = bulletin_res.montant_mutu_patr # F59
    else:
        excel_data_V2[30][5] = 0 # F59

    #ligne 60
    excel_data_V2[31][3] = 20 # D60

    #ligne 61
    excel_data_V2[32][1] = bulletin_res.panier_prix # B61
    excel_data_V2[32][2] = bulletin_res.paniers_sal # C61
    excel_data_V2[32][3] = excel_data_V2[32][1] * excel_data_V2[32][2] # D61
    excel_data_V2[32][5] = excel_data_V2[32][3] # F61

    #ligne 62
    excel_data_V2[33][5] = 0
    for i in range(22,33):
        excel_data_V2[33][5] = excel_data_V2[33][5] + excel_data_V2[i][5] # F62

    #ligne 63
    somme = 0
    for i in range(24,33):
        somme = somme + excel_data_V2[i][3]

    excel_data_V2[34][3] = excel_data_V2[2][3] - excel_data_V2[22][3] + somme # D63

    #ligne 65
    excel_data_V2[36][1] = excel_data_V2[33][5] + excel_data_V2[2][3] # B65

    #ligne 66
    excel_data_V2[37][1] = excel_data_V2[36][1]*12 # B66



    # Formatage des datas (arrondis, %, ...)
    for i in range(7, 28):
        if excel_data_V2[i][2] is not None:
            excel_data_V2[i][2] = excel_data_V2[i][2] * 100

    for i in range(7, 25):
        if excel_data_V2[i][4] is not None:
            excel_data_V2[i][4] = excel_data_V2[i][4] * 100

    for i in range(2,38):
        for j in range(1,6):
            if (excel_data_V2[i][j] is not None) and not isinstance(excel_data_V2[i][j], str):
                if (i== 2) and (j==2):
                    excel_data_V2[i][j] = f"{excel_data_V2[i][j]:.4f}" # attention ça fait que maintenant les chiffres deviennent des strings
                else:
                    excel_data_V2[i][j] = f"{excel_data_V2[i][j]:.2f}" # attention ça fait que maintenant les chiffres deviennent des strings

                if (excel_data_V2[i][j] == f"{0.0:.2f}"):
                    excel_data_V2[i][j] = " "

                excel_data_V2[i][j] = str(excel_data_V2[i][j])

                for p in range(0, len(excel_data_V2[i][j])):
                    if (excel_data_V2[i][j][p] == ".") and (len(excel_data_V2[i][j][:p]) >= 4) and (len(excel_data_V2[i][j][:p]) < 7):
                        u = p - 3
                        excel_data_V2[i][j] = excel_data_V2[i][j][:u] + " " + excel_data_V2[i][j][u:]
                                

    excel_data_V2[34][3] = excel_data_V2[34][3] + " €"
    excel_data_V2[36][1] = excel_data_V2[36][1] + " €"
    excel_data_V2[37][1] = excel_data_V2[37][1] + " €"

    for i in range(7, 28):
        if excel_data_V2[i][2] is not None and excel_data_V2[i][2] != " ":
            excel_data_V2[i][2] = str(excel_data_V2[i][2]) + "%"

    for i in range(7, 25):
        if excel_data_V2[i][4] is not None and excel_data_V2[i][4] != " ":
            excel_data_V2[i][4] = str(excel_data_V2[i][4]) + "%"


    #Suppression des lignes entierement à 0
    # def remove_zero_rows_except_first_column(matrix):
    #     return [row for row in matrix if any(cell != f" " for cell in row[1:])]

    # excel_data_V2 = remove_zero_rows_except_first_column(excel_data_V2)
    # Pour l'instant on laisse en com parce que faut adapter le CSS pour que ça rende bien vu que la c'est sur des chiffres en dur


    # Pass the formatted data to the template
    context = {
        'brut_estime': brut_estime,
        'salaire_net_souhaité': salaire_net_souhaité,
        'bulletin': bulletin_res,
        'excel_data': excel_data_V2
    }

    return render(request, 'simulateur_paie/rendu_Net_to_brut.html', context) 



def simu_net_avec_cout_estime(brut_estime,bulletin_id):

    bulletin = Bulletin_Cout_Vers_Brut.objects.get(pk=bulletin_id)

    excel_data_V2 = [[0 for _ in range(6)] for _ in range(38)]

    # Titres des colonnes
    excel_data_V2[0][0] = "Libellé"
    excel_data_V2[0][1] = "Base"
    excel_data_V2[0][2] = "Taux salarial"
    excel_data_V2[0][3] = "Montant salarial"
    excel_data_V2[0][4] = "Taux patronal"
    excel_data_V2[0][5] = "Montant patronal"

    # Titres des lignes
    excel_data_V2[2][0] = "SALAIRE DE BASE MENSUEL"
    excel_data_V2[3][0] = "VARIABLES DE REMUNERATION"

    excel_data_V2[5][0] = "TOTAL SALAIRE BRUT"

    excel_data_V2[7][0] = "CAR"
    excel_data_V2[8][0] = "CCSS"
    excel_data_V2[9][0] = "Réduction CCSS Gens de maison (33%)"
    excel_data_V2[10][0] = "POLE EMPLOI"
    excel_data_V2[11][0] = "RETRAITE COMPLEMENTAIRE CMRC T1"
    excel_data_V2[12][0] = "RETRAITE COMPLEMENTAIRE CMRC T2"
    excel_data_V2[13][0] = "PREVOYANCE MC TRANCHE A"
    excel_data_V2[14][0] = "PREVOYANCE MC TRANCHE B"
    excel_data_V2[15][0] = "PREVOYANCE FR TRANCHE A"
    excel_data_V2[16][0] = "PREVOYANCE FR TRANCHE B"
    excel_data_V2[17][0] = "ACCIDENT DU TRAVAIL"
    excel_data_V2[18][0] = "TAXE SUR COTISATION AT"
    excel_data_V2[19][0] = "CCPB OUVRIER"
    excel_data_V2[20][0] = "CCPB ETAM/CADRE"

    excel_data_V2[22][0] = "Total cotisations"

    excel_data_V2[24][0] = "Taux de charges"

    excel_data_V2[26][0] = "Non-Soumis"
    excel_data_V2[27][0] = "Indemnité 5% Monégasque"
    excel_data_V2[28][0] = "Remboursement Transport"
    excel_data_V2[29][0] = "Tickets Restaurant"
    excel_data_V2[30][0] = "Mutuelle"
    excel_data_V2[31][0] = "Indemnité Télétravail mensuelle"
    excel_data_V2[32][0] = "Paniers"

    excel_data_V2[34][0] = "Net à payer"

    excel_data_V2[36][0] = "COUT TOTAL MENSUEL EMPLOYEUR"
    excel_data_V2[37][0] = "COUT ANNUEL TOTAL EMPLOYEUR"


    # Taux et plafonds :

    CAR = 6028 # B8
    CAR_part_salariale = 0.0685 # C8
    CAR_part_patronale = 0.0831 # E8

    CCSS = 9600 # B9
    CCSS_part_patronale = 0.1345 # E9

    CCSS_taux_reduit = CCSS # B10
    CCSS_taux_reduit_part_patronale = 0.1340 # E10

    plafond_fracais_tranche_A = 3864 # B11

    SMIC = 11.65 # B12

    Horaire_temps_plein_monegasque = 169 # B13

    plafond_plein_CMRC = 151 # E13

    pole_emploi = plafond_fracais_tranche_A * 4 # B16
    pole_emploi_part_salariale = 0.024 # C16
    pole_emploi_part_patronale = 0.0405 # E16

    CMRC_tranche1 = 3947 # B17
    CMRC_tranche1_part_salariale = 0.0401 # C17
    CMRC_tranche1_part_patronale = 0.0601 # E17

    CMRC_tranche2 = CMRC_tranche1 * 7 # B18
    CMRC_tranche2_part_salariale = 0.0972 # C18
    CMRC_tranche2_part_patronale = 0.1457 # E18

    CCPB_OUVRIER_salariale = 0.004 # C19
    CCPB_OUVRIER_patronale = 0.201 # E19

    CCPB_ETAM_CADRE_salariale = 0.0 # C20
    CCPB_ETAM_CADRE_patronale = 0.205 # E20

    prevoyance_cadre_plafond_monégasque = CCSS # B21
    prevoyance_cadre_plafond_monégasque_patronale = 0.015 # E21

    prevoyance_cadre_plafond_francais =  plafond_fracais_tranche_A # B22
    prevoyance_cadre_plafond_francais_patronale = 0.015 # E22

    taxe_cotisation_AT_patronale = 0.04 # E23

    minimu_garanti = 4.15 # B25



    #ligne 31
    excel_data_V2[2][1] = round(bulletin.heures_mois,2) # B31
    excel_data_V2[2][2] = brut_estime / excel_data_V2[2][1] # C31
    excel_data_V2[2][3] = brut_estime # D31

    #ligne 34
    excel_data_V2[5][3] = excel_data_V2[2][3] # D34

    #ligne 36
    if excel_data_V2[5][3] < CAR:
        excel_data_V2[7][1] = excel_data_V2[5][3] # B36
    else:
        excel_data_V2[7][1] = CAR # B36

    excel_data_V2[7][2] = CAR_part_salariale # C36
    excel_data_V2[7][3] = excel_data_V2[7][1] * excel_data_V2[7][2] # D36
    excel_data_V2[7][4] = CAR_part_patronale # E36
    excel_data_V2[7][5] = excel_data_V2[7][1] * excel_data_V2[7][4] # F36

    #ligne 37
    if bulletin.admin_SAM == "Oui":
        excel_data_V2[8][1] = CCSS # B37
    else:
        excel_data_V2[8][1] = min(excel_data_V2[5][3], CCSS) # B37

    if bulletin.taux_ccss_red == "Oui":
        excel_data_V2[8][4] = CCSS_taux_reduit_part_patronale # E37
    else:
        excel_data_V2[8][4] = CCSS_part_patronale # E37

    excel_data_V2[8][5] = excel_data_V2[8][1] * excel_data_V2[8][4] # F37

    #ligne 38
    if bulletin.Gens_de_Maison == "Oui":
        excel_data_V2[9][1] = -excel_data_V2[8][1] * 0.67 # B38
    else:
        excel_data_V2[9][1] = 0 # B38

    excel_data_V2[9][4] = CCSS_taux_reduit_part_patronale # E38
    excel_data_V2[9][5] = excel_data_V2[9][1] * excel_data_V2[9][4] # F38

    #ligne 39
    H39 = (plafond_fracais_tranche_A / 169 * excel_data_V2[2][1]) * 4 #********************************************************

    if bulletin.Exclu_ass_cho == "Oui":
        excel_data_V2[10][1] = 0 # B39
    else:
        excel_data_V2[10][1] = min(excel_data_V2[5][3], H39) # B39

    excel_data_V2[10][2] = pole_emploi_part_salariale # C39
    excel_data_V2[10][3] = excel_data_V2[10][1] * excel_data_V2[10][2] # D39
    excel_data_V2[10][4] = pole_emploi_part_patronale # E39
    excel_data_V2[10][5] = excel_data_V2[10][1] * excel_data_V2[10][4] # F39

    #ligne 40
    rounded_B31 = round(excel_data_V2[2][1])  # Arrondi de B31 à 0 décimale
    calculation = (CMRC_tranche1 / plafond_plein_CMRC) * rounded_B31 
    H40 = min(CMRC_tranche1, calculation)

    if excel_data_V2[5][3] <= H40:
        excel_data_V2[11][1] = excel_data_V2[5][3] # B40
    else:
        excel_data_V2[11][1] = H40 # B40

    excel_data_V2[11][2] = CMRC_tranche1_part_salariale # C40
    excel_data_V2[11][3] = excel_data_V2[11][1] * excel_data_V2[11][2] # D40
    excel_data_V2[11][4] = CMRC_tranche1_part_patronale # E40
    excel_data_V2[11][5] = excel_data_V2[11][1] * excel_data_V2[11][4] # F40

    #ligne 41
    H41 = H40*7

    if excel_data_V2[5][3] <= (H40 + H41):
        excel_data_V2[12][1] = excel_data_V2[5][3] - excel_data_V2[11][1] # B41
    else:
        excel_data_V2[12][1] = (H40 + H41) - excel_data_V2[11][1] # B41
    
    excel_data_V2[12][2] = CMRC_tranche2_part_salariale  # C41
    excel_data_V2[12][3] = excel_data_V2[12][1] * excel_data_V2[12][2] # D41
    excel_data_V2[12][4] = CMRC_tranche2_part_patronale # E41
    excel_data_V2[12][5] = excel_data_V2[12][1] * excel_data_V2[12][4] # F41

    #ligne 42
    B8 = "Oui" if bulletin.prevoyance else "Non"
    H42 = prevoyance_cadre_plafond_monégasque / Horaire_temps_plein_monegasque * excel_data_V2[2][1]

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[13][1] = 0 # B42
        else:
            excel_data_V2[13][1] = min(H42, excel_data_V2[5][3]) # B42
    else:
        excel_data_V2[13][1] = 0 # B42
    
    excel_data_V2[13][2] = bulletin.taux_prev_T1_sal / 100 # C42
    excel_data_V2[13][3] = excel_data_V2[13][1] * excel_data_V2[13][2]  # D42
    excel_data_V2[13][4] = bulletin.taux_prev_T1_patr / 100 # E42
    excel_data_V2[13][5] = excel_data_V2[13][1] * excel_data_V2[13][4] # F42

    #ligne 43
    H43 = H39 - H42

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[14][1] = 0 # B43
        else:
            if excel_data_V2[5][3] > H42:
                excel_data_V2[14][1] = min(excel_data_V2[5][3]-excel_data_V2[13][1],H42+H43-excel_data_V2[13][1]) # B43
            else:
                excel_data_V2[14][1] = 0 # B43
    else:
        excel_data_V2[14][1] = 0 # B43
    
    excel_data_V2[14][2] =  bulletin.taux_prev_T2_sal / 100 # C43
    excel_data_V2[14][3] =  excel_data_V2[14][1] * excel_data_V2[14][2] # D43
    excel_data_V2[14][4] =  bulletin.taux_prev_T2_patr / 100 # E43
    excel_data_V2[14][5] =  excel_data_V2[14][1] * excel_data_V2[14][4] # F43

    #ligne 44
    H46 = plafond_fracais_tranche_A / Horaire_temps_plein_monegasque * excel_data_V2[2][1]

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[15][1] = min(excel_data_V2[5][3],H46) # B44
        else:
            excel_data_V2[15][1] = 0 # B44
    else:
        excel_data_V2[15][1] = 0 # B44

    excel_data_V2[15][2] = bulletin.taux_prev_T1_sal / 100 # C44
    excel_data_V2[15][3] = excel_data_V2[15][1] * excel_data_V2[15][2]  # D44
    excel_data_V2[15][4] = bulletin.taux_prev_T1_patr / 100 # E44
    excel_data_V2[15][5] = excel_data_V2[15][1] * excel_data_V2[15][4] # F44

    #ligne 45
    H49 = H46*4

    if B8 == "Non":
        excel_data_V2[16][1] = 0 # B45
    else:
        if bulletin.type_plafond != "FR URSSAF":
            excel_data_V2[16][1]  = 0 # B45
        else:
            if excel_data_V2[5][3] > H46: 
                excel_data_V2[16][1] = min(excel_data_V2[5][3] - excel_data_V2[15][1], H49 - excel_data_V2[15][1]) # B45
            else:
                excel_data_V2[16][1] = 0 # B45

    excel_data_V2[16][2] = bulletin.taux_prev_T2_sal / 100 # C45
    excel_data_V2[16][3] = excel_data_V2[16][1] * excel_data_V2[16][2] # D45
    excel_data_V2[16][4] = bulletin.taux_prev_T2_patr / 100 # E45
    excel_data_V2[16][3] = excel_data_V2[16][1] * excel_data_V2[16][4] # F45


    #ligne 46
    excel_data_V2[17][1] =  excel_data_V2[5][3] # B46

    if bulletin.Taux_ATHT > 0:
        excel_data_V2[17][4] = bulletin.Taux_ATHT / 100 # E46
    else:
        excel_data_V2[17][4] = 1 / 100 # E46

    excel_data_V2[17][5] = excel_data_V2[17][1] * excel_data_V2[17][4] # F46

    #ligne 47
    excel_data_V2[18][1] = excel_data_V2[17][5] # B47

    excel_data_V2[18][4] = taxe_cotisation_AT_patronale # E47

    excel_data_V2[18][5] = excel_data_V2[18][1] * excel_data_V2[18][4] # F47

    #ligne 48
    if bulletin.CCPB_ouvrier == "Oui":
        excel_data_V2[19][1] = excel_data_V2[5][3] # B48
    else:
        excel_data_V2[19][1] = 0 # B48

    excel_data_V2[19][2] = CCPB_OUVRIER_salariale # C48

    excel_data_V2[19][3] = excel_data_V2[19][1] * excel_data_V2[19][2] # D48
    excel_data_V2[19][4] = CCPB_OUVRIER_patronale # E48
    excel_data_V2[19][5] = excel_data_V2[19][1] * excel_data_V2[19][4] # F48
    

    #ligne 49

    if bulletin.CCPB_Etam_cadre == "Oui":
        excel_data_V2[20][1] = excel_data_V2[5][3] # B49
    else:
        excel_data_V2[20][1] = 0 # B49

    excel_data_V2[20][2] = CCPB_ETAM_CADRE_salariale # C49 

    excel_data_V2[20][3] = excel_data_V2[20][1] * excel_data_V2[20][2] # D49
    excel_data_V2[20][4] = CCPB_ETAM_CADRE_patronale # E49
    excel_data_V2[20][5] = excel_data_V2[20][1] * excel_data_V2[20][4] # F49


    #ligne 51
    excel_data_V2[22][3] = 0
    for i in range(7,21):
        excel_data_V2[22][3] = excel_data_V2[22][3] + excel_data_V2[i][3] # D51

    excel_data_V2[22][5] = 0
    for i in range(7,21):
        excel_data_V2[22][5] = excel_data_V2[22][5] + excel_data_V2[i][5] # F51

    #ligne 53
    excel_data_V2[24][2] = excel_data_V2[22][3] / excel_data_V2[5][3] # C53

    excel_data_V2[24][4] = excel_data_V2[22][5] / excel_data_V2[5][3] # E53

    #ligne 56
    H50 = SMIC * 1.05

    
    if excel_data_V2[2][2] <= H50:
        excel_data_V2[27][1] = excel_data_V2[5][3] # B56
    else:
        excel_data_V2[27][1] = 0 # B56
    

    excel_data_V2[27][2] = 0.05 # C56

    excel_data_V2[27][3] = excel_data_V2[27][1] * excel_data_V2[27][2] # D56

    excel_data_V2[27][5] = excel_data_V2[27][3] # F56

    #ligne 57
    excel_data_V2[28][3] = bulletin.remboursement_transport # D57

    excel_data_V2[28][5] = excel_data_V2[28][3] # F57

    #ligne 58
    excel_data_V2[29][1] = bulletin.ticket_resto_prix # B58
    excel_data_V2[29][2] = -bulletin.ticket_resto_sal # C58
    excel_data_V2[29][3] = excel_data_V2[29][1] * excel_data_V2[29][2] # D58
    excel_data_V2[29][4] = bulletin.ticket_resto_patr # E58
    excel_data_V2[29][5] = excel_data_V2[29][4] * excel_data_V2[29][1] # F58

    #ligne 59
    B12 = "Oui" if bulletin.mutuelle else "Non"

    if B12 == "Oui":
        excel_data_V2[30][3] = -bulletin.montant_mutu_sal # D59
    else:
        excel_data_V2[30][3] = 0 # D59

    if B12 == "Oui":
        excel_data_V2[30][5] = bulletin.montant_mutu_patr # F59
    else:
        excel_data_V2[30][5] = 0 # F59

    #ligne 60
    excel_data_V2[31][3] = 20 # D60

    #ligne 61
    excel_data_V2[32][1] = bulletin.panier_prix # B61
    excel_data_V2[32][2] = bulletin.paniers_sal # C61
    excel_data_V2[32][3] = excel_data_V2[32][1] * excel_data_V2[32][2] # D61
    excel_data_V2[32][5] = excel_data_V2[32][3] # F61

    #ligne 62
    excel_data_V2[33][5] = 0
    for i in range(22,33):
        excel_data_V2[33][5] = excel_data_V2[33][5] + excel_data_V2[i][5] # F62

    #ligne 63
    somme = 0
    for i in range(24,33):
        somme = somme + excel_data_V2[i][3]

    excel_data_V2[34][3] = excel_data_V2[2][3] - excel_data_V2[22][3] + somme # D63

    #ligne 65
    excel_data_V2[36][1] = excel_data_V2[33][5] + excel_data_V2[2][3] # B65


    excel_data_V2[36][1] = round(excel_data_V2[36][1],2)

    return excel_data_V2[36][1]



def calculer_cout_vers_brut(request, bulletin_id):

    bulletin_res = Bulletin_Cout_Vers_Brut.objects.get(pk=bulletin_id)
    cout_souhaité = bulletin_res.cout_mois # ancien nom : salaire_net_souhaité
    indemnite = bulletin_res.indemnite
    heures = bulletin_res.heures_mois
    H50 = 11.65 * 1.05

    # Hypothèse initiale : cout * 0.714 pour estimer un brut de départ
    brut_estime = cout_souhaité * 0.714
    taux_horaire = brut_estime/heures
    
    if indemnite == "Oui":
        i = 0.01
        while taux_horaire > H50:
            brut_estime =  cout_souhaité * (0.714-i)
            taux_horaire = brut_estime/heures
            i = i + 0.01

    
    max_iter=10000000  # 10 000 000

    for i in range(max_iter):
       
        cout_simulé = simu_net_avec_cout_estime(brut_estime,bulletin_id)
        
        # Comparer le net simulé au net souhaité
        différence = cout_souhaité - cout_simulé
        
        # Si la différence est inférieure à la tolérance, on a trouvé une solution
        if abs(différence) == 0:
            break
        
        # Ajuster le brut estimé en fonction de la différence
        brut_estime += différence * 0.1  # Ajustement itératif léger

    brut_estime = round(brut_estime,1)

    Cout_to_Brut.objects.create(salaire_brut=brut_estime,cout_mois=cout_souhaité, bulletin=bulletin_res)


    excel_data_V2 = [[0 for _ in range(6)] for _ in range(38)]

    # Titres des colonnes
    excel_data_V2[0][0] = "Libellé"
    excel_data_V2[0][1] = "Base"
    excel_data_V2[0][2] = "Taux salarial"
    excel_data_V2[0][3] = "Montant salarial"
    excel_data_V2[0][4] = "Taux patronal"
    excel_data_V2[0][5] = "Montant patronal"

    # Titres des lignes
    excel_data_V2[2][0] = "SALAIRE DE BASE MENSUEL"
    excel_data_V2[3][0] = "VARIABLES DE REMUNERATION"

    excel_data_V2[5][0] = "TOTAL SALAIRE BRUT"

    excel_data_V2[7][0] = "CAR"
    excel_data_V2[8][0] = "CCSS"
    excel_data_V2[9][0] = "Réduction CCSS Gens de maison (33%)"
    excel_data_V2[10][0] = "POLE EMPLOI"
    excel_data_V2[11][0] = "RETRAITE COMPLEMENTAIRE CMRC T1"
    excel_data_V2[12][0] = "RETRAITE COMPLEMENTAIRE CMRC T2"
    excel_data_V2[13][0] = "PREVOYANCE MC TRANCHE A"
    excel_data_V2[14][0] = "PREVOYANCE MC TRANCHE B"
    excel_data_V2[15][0] = "PREVOYANCE FR TRANCHE A"
    excel_data_V2[16][0] = "PREVOYANCE FR TRANCHE B"
    excel_data_V2[17][0] = "ACCIDENT DU TRAVAIL"
    excel_data_V2[18][0] = "TAXE SUR COTISATION AT"
    excel_data_V2[19][0] = "CCPB OUVRIER"
    excel_data_V2[20][0] = "CCPB ETAM/CADRE"

    excel_data_V2[22][0] = "Total cotisations"

    excel_data_V2[24][0] = "Taux de charges"

    excel_data_V2[26][0] = "Non-Soumis"
    excel_data_V2[27][0] = "Indemnité 5% Monégasque"
    excel_data_V2[28][0] = "Remboursement Transport"
    excel_data_V2[29][0] = "Tickets Restaurant"
    excel_data_V2[30][0] = "Mutuelle"
    excel_data_V2[31][0] = "Indemnité Télétravail mensuelle"
    excel_data_V2[32][0] = "Paniers"

    excel_data_V2[34][0] = "Net à payer"

    excel_data_V2[36][0] = "COUT TOTAL MENSUEL EMPLOYEUR"
    excel_data_V2[37][0] = "COUT ANNUEL TOTAL EMPLOYEUR"


    # Taux et plafonds :

    CAR = 6028 # B8
    CAR_part_salariale = 0.0685 # C8
    CAR_part_patronale = 0.0831 # E8

    CCSS = 9600 # B9
    CCSS_part_patronale = 0.1345 # E9

    CCSS_taux_reduit = CCSS # B10
    CCSS_taux_reduit_part_patronale = 0.1340 # E10

    plafond_fracais_tranche_A = 3864 # B11

    SMIC = 11.65 # B12

    Horaire_temps_plein_monegasque = 169 # B13

    plafond_plein_CMRC = 151 # E13

    pole_emploi = plafond_fracais_tranche_A * 4 # B16
    pole_emploi_part_salariale = 0.024 # C16
    pole_emploi_part_patronale = 0.0405 # E16

    CMRC_tranche1 = 3947 # B17
    CMRC_tranche1_part_salariale = 0.0401 # C17
    CMRC_tranche1_part_patronale = 0.0601 # E17

    CMRC_tranche2 = CMRC_tranche1 * 7 # B18
    CMRC_tranche2_part_salariale = 0.0972 # C18
    CMRC_tranche2_part_patronale = 0.1457 # E18

    CCPB_OUVRIER_salariale = 0.004 # C19
    CCPB_OUVRIER_patronale = 0.201 # E19

    CCPB_ETAM_CADRE_salariale = 0.0 # C20
    CCPB_ETAM_CADRE_patronale = 0.205 # E20

    prevoyance_cadre_plafond_monégasque = CCSS # B21
    prevoyance_cadre_plafond_monégasque_patronale = 0.015 # E21

    prevoyance_cadre_plafond_francais =  plafond_fracais_tranche_A # B22
    prevoyance_cadre_plafond_francais_patronale = 0.015 # E22

    taxe_cotisation_AT_patronale = 0.04 # E23

    minimu_garanti = 4.15 # B25



    #ligne 31
    excel_data_V2[2][1] = round(bulletin_res.heures_mois,2) # B31
    excel_data_V2[2][2] = brut_estime / excel_data_V2[2][1] # C31
    excel_data_V2[2][3] = brut_estime # D31

    #ligne 34
    excel_data_V2[5][3] = excel_data_V2[2][3] # D34

    #ligne 36
    if excel_data_V2[5][3] < CAR:
        excel_data_V2[7][1] = excel_data_V2[5][3] # B36
    else:
        excel_data_V2[7][1] = CAR # B36

    excel_data_V2[7][2] = CAR_part_salariale # C36
    excel_data_V2[7][3] = excel_data_V2[7][1] * excel_data_V2[7][2] # D36
    excel_data_V2[7][4] = CAR_part_patronale # E36
    excel_data_V2[7][5] = excel_data_V2[7][1] * excel_data_V2[7][4] # F36

    #ligne 37
    if bulletin_res.admin_SAM == "Oui":
        excel_data_V2[8][1] = CCSS # B37
    else:
        excel_data_V2[8][1] = min(excel_data_V2[5][3], CCSS) # B37

    if bulletin_res.taux_ccss_red == "Oui":
        excel_data_V2[8][4] = CCSS_taux_reduit_part_patronale # E37
    else:
        excel_data_V2[8][4] = CCSS_part_patronale # E37

    excel_data_V2[8][5] = excel_data_V2[8][1] * excel_data_V2[8][4] # F37

    #ligne 38
    if bulletin_res.Gens_de_Maison == "Oui":
        excel_data_V2[9][1] = -excel_data_V2[8][1] * 0.67 # B38
    else:
        excel_data_V2[9][1] = 0 # B38

    excel_data_V2[9][4] = CCSS_taux_reduit_part_patronale # E38
    excel_data_V2[9][5] = excel_data_V2[9][1] * excel_data_V2[9][4] # F38

    #ligne 39
    H39 = (plafond_fracais_tranche_A / 169 * excel_data_V2[2][1]) * 4 #********************************************************

    if bulletin_res.Exclu_ass_cho == "Oui":
        excel_data_V2[10][1] = 0 # B39
    else:
        excel_data_V2[10][1] = min(excel_data_V2[5][3], H39) # B39

    excel_data_V2[10][2] = pole_emploi_part_salariale # C39
    excel_data_V2[10][3] = excel_data_V2[10][1] * excel_data_V2[10][2] # D39
    excel_data_V2[10][4] = pole_emploi_part_patronale # E39
    excel_data_V2[10][5] = excel_data_V2[10][1] * excel_data_V2[10][4] # F39

    #ligne 40
    rounded_B31 = round(excel_data_V2[2][1])  # Arrondi de B31 à 0 décimale
    calculation = (CMRC_tranche1 / plafond_plein_CMRC) * rounded_B31 
    H40 = min(CMRC_tranche1, calculation)

    if excel_data_V2[5][3] <= H40:
        excel_data_V2[11][1] = excel_data_V2[5][3] # B40
    else:
        excel_data_V2[11][1] = H40 # B40

    excel_data_V2[11][2] = CMRC_tranche1_part_salariale # C40
    excel_data_V2[11][3] = excel_data_V2[11][1] * excel_data_V2[11][2] # D40
    excel_data_V2[11][4] = CMRC_tranche1_part_patronale # E40
    excel_data_V2[11][5] = excel_data_V2[11][1] * excel_data_V2[11][4] # F40

    #ligne 41
    H41 = H40*7

    if excel_data_V2[5][3] <= (H40 + H41):
        excel_data_V2[12][1] = excel_data_V2[5][3] - excel_data_V2[11][1] # B41
    else:
        excel_data_V2[12][1] = (H40 + H41) - excel_data_V2[11][1] # B41
    
    excel_data_V2[12][2] = CMRC_tranche2_part_salariale  # C41
    excel_data_V2[12][3] = excel_data_V2[12][1] * excel_data_V2[12][2] # D41
    excel_data_V2[12][4] = CMRC_tranche2_part_patronale # E41
    excel_data_V2[12][5] = excel_data_V2[12][1] * excel_data_V2[12][4] # F41

    #ligne 42
    B8 = "Oui" if bulletin_res.prevoyance else "Non"
    H42 = prevoyance_cadre_plafond_monégasque / Horaire_temps_plein_monegasque * excel_data_V2[2][1]

    if B8 == "Oui":
        if bulletin_res.type_plafond == "FR URSSAF":
            excel_data_V2[13][1] = 0 # B42
        else:
            excel_data_V2[13][1] = min(H42, excel_data_V2[5][3]) # B42
    else:
        excel_data_V2[13][1] = 0 # B42
    
    excel_data_V2[13][2] = bulletin_res.taux_prev_T1_sal / 100 # C42
    excel_data_V2[13][3] = excel_data_V2[13][1] * excel_data_V2[13][2]  # D42
    excel_data_V2[13][4] = bulletin_res.taux_prev_T1_patr / 100 # E42
    excel_data_V2[13][5] = excel_data_V2[13][1] * excel_data_V2[13][4] # F42

    #ligne 43
    H43 = H39 - H42

    if B8 == "Oui":
        if bulletin_res.type_plafond == "FR URSSAF":
            excel_data_V2[14][1] = 0 # B43
        else:
            if excel_data_V2[5][3] > H42:
                excel_data_V2[14][1] = min(excel_data_V2[5][3]-excel_data_V2[13][1],H42+H43-excel_data_V2[13][1]) # B43
            else:
                excel_data_V2[14][1] = 0 # B43
    else:
        excel_data_V2[14][1] = 0 # B43
    
    excel_data_V2[14][2] =  bulletin_res.taux_prev_T2_sal / 100 # C43
    excel_data_V2[14][3] =  excel_data_V2[14][1] * excel_data_V2[14][2] # D43
    excel_data_V2[14][4] =  bulletin_res.taux_prev_T2_patr / 100 # E43
    excel_data_V2[14][5] =  excel_data_V2[14][1] * excel_data_V2[14][4] # F43

    #ligne 44
    H46 = plafond_fracais_tranche_A / Horaire_temps_plein_monegasque * excel_data_V2[2][1]

    if B8 == "Oui":
        if bulletin_res.type_plafond == "FR URSSAF":
            excel_data_V2[15][1] = min(excel_data_V2[5][3],H46) # B44
        else:
            excel_data_V2[15][1] = 0 # B44
    else:
        excel_data_V2[15][1] = 0 # B44

    excel_data_V2[15][2] = bulletin_res.taux_prev_T1_sal / 100 # C44
    excel_data_V2[15][3] = excel_data_V2[15][1] * excel_data_V2[15][2]  # D44
    excel_data_V2[15][4] = bulletin_res.taux_prev_T1_patr / 100 # E44
    excel_data_V2[15][5] = excel_data_V2[15][1] * excel_data_V2[15][4] # F44

    #ligne 45
    H49 = H46*4

    if B8 == "Non":
        excel_data_V2[16][1] = 0 # B45
    else:
        if bulletin_res.type_plafond != "FR URSSAF":
            excel_data_V2[16][1]  = 0 # B45
        else:
            if excel_data_V2[5][3] > H46: 
                excel_data_V2[16][1] = min(excel_data_V2[5][3] - excel_data_V2[15][1], H49 - excel_data_V2[15][1]) # B45
            else:
                excel_data_V2[16][1] = 0 # B45

    excel_data_V2[16][2] = bulletin_res.taux_prev_T2_sal / 100 # C45
    excel_data_V2[16][3] = excel_data_V2[16][1] * excel_data_V2[16][2] # D45
    excel_data_V2[16][4] = bulletin_res.taux_prev_T2_patr / 100 # E45
    excel_data_V2[16][3] = excel_data_V2[16][1] * excel_data_V2[16][4] # F45


    #ligne 46
    excel_data_V2[17][1] =  excel_data_V2[5][3] # B46

    if bulletin_res.Taux_ATHT > 0:
        excel_data_V2[17][4] = bulletin_res.Taux_ATHT / 100 # E46
    else:
        excel_data_V2[17][4] = 1 / 100 # E46

    excel_data_V2[17][5] = excel_data_V2[17][1] * excel_data_V2[17][4] # F46

    #ligne 47
    excel_data_V2[18][1] = excel_data_V2[17][5] # B47

    excel_data_V2[18][4] = taxe_cotisation_AT_patronale # E47

    excel_data_V2[18][5] = excel_data_V2[18][1] * excel_data_V2[18][4] # F47

    #ligne 48
    if bulletin_res.CCPB_ouvrier == "Oui":
        excel_data_V2[19][1] = excel_data_V2[5][3] # B48
    else:
        excel_data_V2[19][1] = 0 # B48

    excel_data_V2[19][2] = CCPB_OUVRIER_salariale # C48

    excel_data_V2[19][3] = excel_data_V2[19][1] * excel_data_V2[19][2] # D48
    excel_data_V2[19][4] = CCPB_OUVRIER_patronale # E48
    excel_data_V2[19][5] = excel_data_V2[19][1] * excel_data_V2[19][4] # F48
    

    #ligne 49

    if bulletin_res.CCPB_Etam_cadre == "Oui":
        excel_data_V2[20][1] = excel_data_V2[5][3] # B49
    else:
        excel_data_V2[20][1] = 0 # B49

    excel_data_V2[20][2] = CCPB_ETAM_CADRE_salariale # C49 

    excel_data_V2[20][3] = excel_data_V2[20][1] * excel_data_V2[20][2] # D49
    excel_data_V2[20][4] = CCPB_ETAM_CADRE_patronale # E49
    excel_data_V2[20][5] = excel_data_V2[20][1] * excel_data_V2[20][4] # F49


    #ligne 51
    excel_data_V2[22][3] = 0
    for i in range(7,21):
        excel_data_V2[22][3] = excel_data_V2[22][3] + excel_data_V2[i][3] # D51

    excel_data_V2[22][5] = 0
    for i in range(7,21):
        excel_data_V2[22][5] = excel_data_V2[22][5] + excel_data_V2[i][5] # F51

    #ligne 53
    excel_data_V2[24][2] = excel_data_V2[22][3] / excel_data_V2[5][3] # C53

    excel_data_V2[24][4] = excel_data_V2[22][5] / excel_data_V2[5][3] # E53

    #ligne 56
    H50 = SMIC * 1.05

    if excel_data_V2[2][2] <= H50:
        excel_data_V2[27][1] = excel_data_V2[5][3] # B56
    else:
        excel_data_V2[27][1] = 0 # B56

    excel_data_V2[27][2] = 0.05 # C56

    excel_data_V2[27][3] = excel_data_V2[27][1] * excel_data_V2[27][2] # D56

    excel_data_V2[27][5] = excel_data_V2[27][3] # F56

    #ligne 57
    excel_data_V2[28][3] = bulletin_res.remboursement_transport # D57

    excel_data_V2[28][5] = excel_data_V2[28][3] # F57

    #ligne 58
    excel_data_V2[29][1] = bulletin_res.ticket_resto_prix # B58
    excel_data_V2[29][2] = -bulletin_res.ticket_resto_sal # C58
    excel_data_V2[29][3] = excel_data_V2[29][1] * excel_data_V2[29][2] # D58
    excel_data_V2[29][4] = bulletin_res.ticket_resto_patr # E58
    excel_data_V2[29][5] = excel_data_V2[29][4] * excel_data_V2[29][1] # F58

    #ligne 59
    B12 = "Oui" if bulletin_res.mutuelle else "Non"

    if B12 == "Oui":
        excel_data_V2[30][3] = -bulletin_res.montant_mutu_sal # D59
    else:
        excel_data_V2[30][3] = 0 # D59

    if B12 == "Oui":
        excel_data_V2[30][5] = bulletin_res.montant_mutu_patr # F59
    else:
        excel_data_V2[30][5] = 0 # F59

    #ligne 60
    excel_data_V2[31][3] = 20 # D60

    #ligne 61
    excel_data_V2[32][1] = bulletin_res.panier_prix # B61
    excel_data_V2[32][2] = bulletin_res.paniers_sal # C61
    excel_data_V2[32][3] = excel_data_V2[32][1] * excel_data_V2[32][2] # D61
    excel_data_V2[32][5] = excel_data_V2[32][3] # F61

    #ligne 62
    excel_data_V2[33][5] = 0
    for i in range(22,33):
        excel_data_V2[33][5] = excel_data_V2[33][5] + excel_data_V2[i][5] # F62

    #ligne 63
    somme = 0
    for i in range(24,33):
        somme = somme + excel_data_V2[i][3]

    excel_data_V2[34][3] = excel_data_V2[2][3] - excel_data_V2[22][3] + somme # D63

    #ligne 65
    excel_data_V2[36][1] = excel_data_V2[33][5] + excel_data_V2[2][3] # B65

    #ligne 66
    excel_data_V2[37][1] = excel_data_V2[36][1]*12 # B66



    # Formatage des datas (arrondis, %, ...)
    for i in range(7, 28):
        if excel_data_V2[i][2] is not None:
            excel_data_V2[i][2] = excel_data_V2[i][2] * 100

    for i in range(7, 25):
        if excel_data_V2[i][4] is not None:
            excel_data_V2[i][4] = excel_data_V2[i][4] * 100

    for i in range(2,38):
        for j in range(1,6):
            if (excel_data_V2[i][j] is not None) and not isinstance(excel_data_V2[i][j], str):
                if (i== 2) and (j==2):
                    excel_data_V2[i][j] = f"{excel_data_V2[i][j]:.4f}" # attention ça fait que maintenant les chiffres deviennent des strings
                else:
                    excel_data_V2[i][j] = f"{excel_data_V2[i][j]:.2f}" # attention ça fait que maintenant les chiffres deviennent des strings

                if (excel_data_V2[i][j] == f"{0.0:.2f}"):
                    excel_data_V2[i][j] = " "

                excel_data_V2[i][j] = str(excel_data_V2[i][j])

                for p in range(0, len(excel_data_V2[i][j])):
                    if (excel_data_V2[i][j][p] == ".") and (len(excel_data_V2[i][j][:p]) >= 4) and (len(excel_data_V2[i][j][:p]) < 7):
                        u = p - 3
                        excel_data_V2[i][j] = excel_data_V2[i][j][:u] + " " + excel_data_V2[i][j][u:]
                                

    excel_data_V2[34][3] = excel_data_V2[34][3] + " €"
    excel_data_V2[36][1] = excel_data_V2[36][1] + " €"
    excel_data_V2[37][1] = excel_data_V2[37][1] + " €"

    for i in range(7, 28):
        if excel_data_V2[i][2] is not None and excel_data_V2[i][2] != " ":
            excel_data_V2[i][2] = str(excel_data_V2[i][2]) + "%"

    for i in range(7, 25):
        if excel_data_V2[i][4] is not None and excel_data_V2[i][4] != " ":
            excel_data_V2[i][4] = str(excel_data_V2[i][4]) + "%"


    #Suppression des lignes entierement à 0
    # def remove_zero_rows_except_first_column(matrix):
    #     return [row for row in matrix if any(cell != f" " for cell in row[1:])]

    # excel_data_V2 = remove_zero_rows_except_first_column(excel_data_V2)
    # Pour l'instant on laisse en com parce que faut adapter le CSS pour que ça rende bien vu que la c'est sur des chiffres en dur


    # Pass the formatted data to the template
    context = {
        'brut_estime': brut_estime,
        'cout_souhaité': cout_souhaité,
        'bulletin': bulletin_res,
        'excel_data': excel_data_V2
    }

    return render(request, 'simulateur_paie/rendu_Cout_to_brut.html', context) # A changer *************************************************





def telecharger_tableau_pdf(request,bulletin_id):
    # Create the HttpResponse object with the appropriate PDF headers.
    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = 'attachment; filename="bulletin_paie.pdf"'

    # Create the PDF object, using the response object as its "file."
    pdf_canvas = canvas.Canvas(response, pagesize=A4)  # Avoid using 'p'
    
    # Get A4 dimensions
    page_width, page_height = A4

    # Title
    pdf_canvas.setFont("Helvetica", 12)
    pdf_canvas.drawString(200, page_height - 40, "Bulletin de Paie")

    bulletin = Bulletin.objects.get(pk=bulletin_id)

    # Simulating the payroll table (excel_data_V2)
    excel_data_V2 = [[0 for _ in range(6)] for _ in range(38)]

    # Titres des colonnes
    excel_data_V2[0][0] = "Libellé"
    excel_data_V2[0][1] = "Base"
    excel_data_V2[0][2] = "Taux salarial"
    excel_data_V2[0][3] = "Montant salarial"
    excel_data_V2[0][4] = "Taux patronal"
    excel_data_V2[0][5] = "Montant patronal"

    # Titres des lignes
    excel_data_V2[2][0] = "SALAIRE DE BASE MENSUEL"
    excel_data_V2[3][0] = "VARIABLES DE REMUNERATION"

    excel_data_V2[5][0] = "TOTAL SALAIRE BRUT"

    excel_data_V2[7][0] = "CAR"
    excel_data_V2[8][0] = "CCSS"
    excel_data_V2[9][0] = "Réduction CCSS Gens de maison (33%)"
    excel_data_V2[10][0] = "POLE EMPLOI"
    excel_data_V2[11][0] = "RETRAITE COMPLEMENTAIRE CMRC T1"
    excel_data_V2[12][0] = "RETRAITE COMPLEMENTAIRE CMRC T2"
    excel_data_V2[13][0] = "PREVOYANCE MC TRANCHE A"
    excel_data_V2[14][0] = "PREVOYANCE MC TRANCHE B"
    excel_data_V2[15][0] = "PREVOYANCE FR TRANCHE A"
    excel_data_V2[16][0] = "PREVOYANCE FR TRANCHE B"
    excel_data_V2[17][0] = "ACCIDENT DU TRAVAIL"
    excel_data_V2[18][0] = "TAXE SUR COTISATION AT"
    excel_data_V2[19][0] = "CCPB OUVRIER"
    excel_data_V2[20][0] = "CCPB ETAM/CADRE"

    excel_data_V2[22][0] = "Total cotisations"

    excel_data_V2[24][0] = "Taux de charges"

    excel_data_V2[26][0] = "Non-Soumis"
    excel_data_V2[27][0] = "Indemnité 5% Monégasque"
    excel_data_V2[28][0] = "Remboursement Transport"
    excel_data_V2[29][0] = "Tickets Restaurant"
    excel_data_V2[30][0] = "Mutuelle"
    excel_data_V2[31][0] = "Indemnité Télétravail mensuelle"
    excel_data_V2[32][0] = "Paniers"

    excel_data_V2[34][0] = "Net à payer"

    excel_data_V2[36][0] = "COUT TOTAL MENSUEL EMPLOYEUR"
    excel_data_V2[37][0] = "COUT ANNUEL TOTAL EMPLOYEUR"


    # Taux et plafonds :

    CAR = 6028 # B8
    CAR_part_salariale = 0.0685 # C8
    CAR_part_patronale = 0.0831 # E8

    CCSS = 9600 # B9
    CCSS_part_patronale = 0.1345 # E9

    CCSS_taux_reduit = CCSS # B10
    CCSS_taux_reduit_part_patronale = 0.1340 # E10

    plafond_fracais_tranche_A = 3864 # B11

    SMIC = 11.65 # B12

    Horaire_temps_plein_monegasque = 169 # B13

    plafond_plein_CMRC = 151 # E13

    pole_emploi = plafond_fracais_tranche_A * 4 # B16
    pole_emploi_part_salariale = 0.024 # C16
    pole_emploi_part_patronale = 0.0405 # E16

    CMRC_tranche1 = 3947 # B17
    CMRC_tranche1_part_salariale = 0.0401 # C17
    CMRC_tranche1_part_patronale = 0.0601 # E17

    CMRC_tranche2 = CMRC_tranche1 * 7 # B18
    CMRC_tranche2_part_salariale = 0.0972 # C18
    CMRC_tranche2_part_patronale = 0.1457 # E18

    CCPB_OUVRIER_salariale = 0.004 # C19
    CCPB_OUVRIER_patronale = 0.201 # E19

    CCPB_ETAM_CADRE_salariale = 0.0 # C20
    CCPB_ETAM_CADRE_patronale = 0.205 # E20

    prevoyance_cadre_plafond_monégasque = CCSS # B21
    prevoyance_cadre_plafond_monégasque_patronale = 0.015 # E21

    prevoyance_cadre_plafond_francais =  plafond_fracais_tranche_A # B22
    prevoyance_cadre_plafond_francais_patronale = 0.015 # E22

    taxe_cotisation_AT_patronale = 0.04 # E23

    minimu_garanti = 4.15 # B25



    #ligne 31
    excel_data_V2[2][1] = round(bulletin.heures_mois,2) # B31
    excel_data_V2[2][2] = bulletin.salaire_brut_mois / excel_data_V2[2][1] # C31
    excel_data_V2[2][3] = bulletin.salaire_brut_mois # D31

    #ligne 34
    excel_data_V2[5][3] = excel_data_V2[2][3] # D34

    #ligne 36
    if excel_data_V2[5][3] < CAR:
        excel_data_V2[7][1] = excel_data_V2[5][3] # B36
    else:
        excel_data_V2[7][1] = CAR # B36

    excel_data_V2[7][2] = CAR_part_salariale # C36
    excel_data_V2[7][3] = excel_data_V2[7][1] * excel_data_V2[7][2] # D36
    excel_data_V2[7][4] = CAR_part_patronale # E36
    excel_data_V2[7][5] = excel_data_V2[7][1] * excel_data_V2[7][4] # F36

    #ligne 37
    if bulletin.admin_SAM == "Oui":
        excel_data_V2[8][1] = CCSS # B37
    else:
        excel_data_V2[8][1] = min(excel_data_V2[5][3], CCSS) # B37

    if bulletin.taux_ccss_red == "Oui":
        excel_data_V2[8][4] = CCSS_taux_reduit_part_patronale # E37
    else:
        excel_data_V2[8][4] = CCSS_part_patronale # E37

    excel_data_V2[8][5] = excel_data_V2[8][1] * excel_data_V2[8][4] # F37

    #ligne 38
    if bulletin.Gens_de_Maison == "Oui":
        excel_data_V2[9][1] = -excel_data_V2[8][1] * 0.67 # B38
    else:
        excel_data_V2[9][1] = 0 # B38

    excel_data_V2[9][4] = CCSS_taux_reduit_part_patronale # E38
    excel_data_V2[9][5] = excel_data_V2[9][1] * excel_data_V2[9][4] # F38

    #ligne 39
    H39 = (plafond_fracais_tranche_A / 169 * excel_data_V2[2][1]) * 4 #********************************************************

    if bulletin.Exclu_ass_cho == "Oui":
        excel_data_V2[10][1] = 0 # B39
    else:
        excel_data_V2[10][1] = min(excel_data_V2[5][3], H39) # B39

    excel_data_V2[10][2] = pole_emploi_part_salariale # C39
    excel_data_V2[10][3] = excel_data_V2[10][1] * excel_data_V2[10][2] # D39
    excel_data_V2[10][4] = pole_emploi_part_patronale # E39
    excel_data_V2[10][5] = excel_data_V2[10][1] * excel_data_V2[10][4] # F39

    #ligne 40
    rounded_B31 = round(excel_data_V2[2][1])  # Arrondi de B31 à 0 décimale
    calculation = (CMRC_tranche1 / plafond_plein_CMRC) * rounded_B31 
    H40 = min(CMRC_tranche1, calculation)

    if excel_data_V2[5][3] <= H40:
        excel_data_V2[11][1] = excel_data_V2[5][3] # B40
    else:
        excel_data_V2[11][1] = H40 # B40

    excel_data_V2[11][2] = CMRC_tranche1_part_salariale # C40
    excel_data_V2[11][3] = excel_data_V2[11][1] * excel_data_V2[11][2] # D40
    excel_data_V2[11][4] = CMRC_tranche1_part_patronale # E40
    excel_data_V2[11][5] = excel_data_V2[11][1] * excel_data_V2[11][4] # F40

    #ligne 41
    H41 = H40*7

    if excel_data_V2[5][3] <= (H40 + H41):
        excel_data_V2[12][1] = excel_data_V2[5][3] - excel_data_V2[11][1] # B41
    else:
        excel_data_V2[12][1] = (H40 + H41) - excel_data_V2[11][1] # B41
    
    excel_data_V2[12][2] = CMRC_tranche2_part_salariale  # C41
    excel_data_V2[12][3] = excel_data_V2[12][1] * excel_data_V2[12][2] # D41
    excel_data_V2[12][4] = CMRC_tranche2_part_patronale # E41
    excel_data_V2[12][5] = excel_data_V2[12][1] * excel_data_V2[12][4] # F41

    #ligne 42
    B8 = "Oui" if bulletin.prevoyance else "Non"
    H42 = prevoyance_cadre_plafond_monégasque / Horaire_temps_plein_monegasque * excel_data_V2[2][1]

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[13][1] = 0 # B42
        else:
            excel_data_V2[13][1] = min(H42, excel_data_V2[5][3]) # B42
    else:
        excel_data_V2[13][1] = 0 # B42
    
    excel_data_V2[13][2] = bulletin.taux_prev_T1_sal / 100 # C42
    excel_data_V2[13][3] = excel_data_V2[13][1] * excel_data_V2[13][2]  # D42
    excel_data_V2[13][4] = bulletin.taux_prev_T1_patr / 100 # E42
    excel_data_V2[13][5] = excel_data_V2[13][1] * excel_data_V2[13][4] # F42

    #ligne 43
    H43 = H39 - H42

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[14][1] = 0 # B43
        else:
            if excel_data_V2[5][3] > H42:
                excel_data_V2[14][1] = min(excel_data_V2[5][3]-excel_data_V2[13][1],H42+H43-excel_data_V2[13][1]) # B43
            else:
                excel_data_V2[14][1] = 0 # B43
    else:
        excel_data_V2[14][1] = 0 # B43
    
    excel_data_V2[14][2] =  bulletin.taux_prev_T2_sal / 100 # C43
    excel_data_V2[14][3] =  excel_data_V2[14][1] * excel_data_V2[14][2] # D43
    excel_data_V2[14][4] =  bulletin.taux_prev_T2_patr / 100 # E43
    excel_data_V2[14][5] =  excel_data_V2[14][1] * excel_data_V2[14][4] # F43

    #ligne 44
    H46 = plafond_fracais_tranche_A / Horaire_temps_plein_monegasque * excel_data_V2[2][1]

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[15][1] = min(excel_data_V2[5][3],H46) # B44
        else:
            excel_data_V2[15][1] = 0 # B44
    else:
        excel_data_V2[15][1] = 0 # B44

    excel_data_V2[15][2] = bulletin.taux_prev_T1_sal / 100 # C44
    excel_data_V2[15][3] = excel_data_V2[15][1] * excel_data_V2[15][2]  # D44
    excel_data_V2[15][4] = bulletin.taux_prev_T1_patr / 100 # E44
    excel_data_V2[15][5] = excel_data_V2[15][1] * excel_data_V2[15][4] # F44

    #ligne 45
    H49 = H46*4

    if B8 == "Non":
        excel_data_V2[16][1] = 0 # B45
    else:
        if bulletin.type_plafond != "FR URSSAF":
            excel_data_V2[16][1]  = 0 # B45
        else:
            if excel_data_V2[5][3] > H46: 
                excel_data_V2[16][1] = min(excel_data_V2[5][3] - excel_data_V2[15][1], H49 - excel_data_V2[15][1]) # B45
            else:
                excel_data_V2[16][1] = 0 # B45

    excel_data_V2[16][2] = bulletin.taux_prev_T2_sal / 100 # C45
    excel_data_V2[16][3] = excel_data_V2[16][1] * excel_data_V2[16][2] # D45
    excel_data_V2[16][4] = bulletin.taux_prev_T2_patr / 100 # E45
    excel_data_V2[16][3] = excel_data_V2[16][1] * excel_data_V2[16][4] # F45


    #ligne 46
    excel_data_V2[17][1] =  excel_data_V2[5][3] # B46

    if bulletin.Taux_ATHT > 0:
        excel_data_V2[17][4] = bulletin.Taux_ATHT / 100 # E46
    else:
        excel_data_V2[17][4] = 1 / 100 # E46

    excel_data_V2[17][5] = excel_data_V2[17][1] * excel_data_V2[17][4] # F46

    #ligne 47
    excel_data_V2[18][1] = excel_data_V2[17][5] # B47

    excel_data_V2[18][4] = taxe_cotisation_AT_patronale # E47

    excel_data_V2[18][5] = excel_data_V2[18][1] * excel_data_V2[18][4] # F47

    #ligne 48
    if bulletin.CCPB_ouvrier == "Oui":
        excel_data_V2[19][1] = excel_data_V2[5][3] # B48
    else:
        excel_data_V2[19][1] = 0 # B48

    excel_data_V2[19][2] = CCPB_OUVRIER_salariale # C48

    excel_data_V2[19][3] = excel_data_V2[19][1] * excel_data_V2[19][2] # D48
    excel_data_V2[19][4] = CCPB_OUVRIER_patronale # E48
    excel_data_V2[19][5] = excel_data_V2[19][1] * excel_data_V2[19][4] # F48
    

    #ligne 49

    if bulletin.CCPB_Etam_cadre == "Oui":
        excel_data_V2[20][1] = excel_data_V2[5][3] # B49
    else:
        excel_data_V2[20][1] = 0 # B49

    excel_data_V2[20][2] = CCPB_ETAM_CADRE_salariale # C49 

    excel_data_V2[20][3] = excel_data_V2[20][1] * excel_data_V2[20][2] # D49
    excel_data_V2[20][4] = CCPB_ETAM_CADRE_patronale # E49
    excel_data_V2[20][5] = excel_data_V2[20][1] * excel_data_V2[20][4] # F49


    #ligne 51
    excel_data_V2[22][3] = 0
    for i in range(7,21):
        excel_data_V2[22][3] = excel_data_V2[22][3] + excel_data_V2[i][3] # D51

    excel_data_V2[22][5] = 0
    for i in range(7,21):
        excel_data_V2[22][5] = excel_data_V2[22][5] + excel_data_V2[i][5] # F51

    #ligne 53
    excel_data_V2[24][2] = excel_data_V2[22][3] / excel_data_V2[5][3] # C53

    excel_data_V2[24][4] = excel_data_V2[22][5] / excel_data_V2[5][3] # E53

    #ligne 56
    H50 = SMIC * 1.05

    if excel_data_V2[2][2] <= H50:
        excel_data_V2[27][1] = excel_data_V2[5][3] # B56
    else:
        excel_data_V2[27][1] = 0 # B56

    excel_data_V2[27][2] = 0.05 # C56

    excel_data_V2[27][3] = excel_data_V2[27][1] * excel_data_V2[27][2] # D56

    excel_data_V2[27][5] = excel_data_V2[27][3] # F56

    #ligne 57
    excel_data_V2[28][3] = bulletin.remboursement_transport # D57

    excel_data_V2[28][5] = excel_data_V2[28][3] # F57

    #ligne 58
    excel_data_V2[29][1] = bulletin.ticket_resto_prix # B58
    excel_data_V2[29][2] = -bulletin.ticket_resto_sal # C58
    excel_data_V2[29][3] = excel_data_V2[29][1] * excel_data_V2[29][2] # D58
    excel_data_V2[29][4] = bulletin.ticket_resto_patr # E58
    excel_data_V2[29][5] = excel_data_V2[29][4] * excel_data_V2[29][1] # F58

    #ligne 59
    B12 = "Oui" if bulletin.mutuelle else "Non"

    if B12 == "Oui":
        excel_data_V2[30][3] = -bulletin.montant_mutu_sal # D59
    else:
        excel_data_V2[30][3] = 0 # D59

    if B12 == "Oui":
        excel_data_V2[30][5] = bulletin.montant_mutu_patr # F59
    else:
        excel_data_V2[30][5] = 0 # F59

    #ligne 60
    excel_data_V2[31][3] = 20 # D60

    #ligne 61
    excel_data_V2[32][1] = bulletin.panier_prix # B61
    excel_data_V2[32][2] = bulletin.paniers_sal # C61
    excel_data_V2[32][3] = excel_data_V2[32][1] * excel_data_V2[32][2] # D61
    excel_data_V2[32][5] = excel_data_V2[32][3] # F61

    #ligne 62
    excel_data_V2[33][5] = 0
    for i in range(22,33):
        excel_data_V2[33][5] = excel_data_V2[33][5] + excel_data_V2[i][5] # F62

    #ligne 63
    somme = 0
    for i in range(24,33):
        somme = somme + excel_data_V2[i][3]

    excel_data_V2[34][3] = excel_data_V2[2][3] - excel_data_V2[22][3] + somme # D63

    #ligne 65
    excel_data_V2[36][1] = excel_data_V2[33][5] + excel_data_V2[2][3] # B65

    #ligne 66
    excel_data_V2[37][1] = excel_data_V2[36][1]*12 # B66



    # Formatage des datas (arrondis, %, ...)
    for i in range(7, 28):
        if excel_data_V2[i][2] is not None:
            excel_data_V2[i][2] = excel_data_V2[i][2] * 100

    for i in range(7, 25):
        if excel_data_V2[i][4] is not None:
            excel_data_V2[i][4] = excel_data_V2[i][4] * 100

    for i in range(2,38):
        for j in range(1,6):
            if (excel_data_V2[i][j] is not None) and not isinstance(excel_data_V2[i][j], str):
                if (i== 2) and (j==2):
                    excel_data_V2[i][j] = f"{excel_data_V2[i][j]:.4f}" # attention ça fait que maintenant les chiffres deviennent des strings
                else:
                    excel_data_V2[i][j] = f"{excel_data_V2[i][j]:.2f}" # attention ça fait que maintenant les chiffres deviennent des strings

                if (excel_data_V2[i][j] == f"{0.0:.2f}"):
                    excel_data_V2[i][j] = " "

                excel_data_V2[i][j] = str(excel_data_V2[i][j])

                for p in range(0, len(excel_data_V2[i][j])):
                    if (excel_data_V2[i][j][p] == ".") and (len(excel_data_V2[i][j][:p]) >= 4) and (len(excel_data_V2[i][j][:p]) < 7):
                        u = p - 3
                        excel_data_V2[i][j] = excel_data_V2[i][j][:u] + " " + excel_data_V2[i][j][u:]
                                

    excel_data_V2[34][3] = excel_data_V2[34][3] + " €"
    excel_data_V2[36][1] = excel_data_V2[36][1] + " €"
    excel_data_V2[37][1] = excel_data_V2[37][1] + " €"

    for i in range(7, 28):
        if excel_data_V2[i][2] is not None and excel_data_V2[i][2] != " ":
            excel_data_V2[i][2] = str(excel_data_V2[i][2]) + "%"

    for i in range(7, 25):
        if excel_data_V2[i][4] is not None and excel_data_V2[i][4] != " ":
            excel_data_V2[i][4] = str(excel_data_V2[i][4]) + "%"

    # Remplacement des 0 par une chaîne vide
    excel_data_V2 = [[" " if element == 0 else element for element in sous_liste] for sous_liste in excel_data_V2]


    # Font for table
    font_name = "Helvetica"
    font_size = 10
    pdf_canvas.setFont(font_name, font_size)

    # Measure the maximum width of each column
    num_columns = len(excel_data_V2[0])
    column_widths = [0] * num_columns  # Store the max width for each column
    
    for row in excel_data_V2:
        for i, item in enumerate(row):
            # Measure width of the text in each column using reportlab's stringWidth method
            text_width = pdf_canvas.stringWidth(str(item), font_name, font_size)
            if text_width > column_widths[i]:
                column_widths[i] = text_width  # Add some padding for readability

    # Check if the total width exceeds the page width and adjust scaling if necessary
    total_table_width = sum(column_widths)
    if total_table_width > page_width - 100:  # Consider margins
        scaling_factor = (page_width - 100) / total_table_width
        column_widths = [width * scaling_factor for width in column_widths]

    # Adjusted layout settings
    x_offset = 50
    y_offset = page_height - 100  # Adjust the starting y position
    cell_height = 20  # Height of each cell

    # Drawing the table with borders
    for row in excel_data_V2:
        current_x = x_offset
        for i, item in enumerate(row):
            pdf_canvas.drawString(current_x + 2, y_offset + 5, str(item))  # Add some padding for text
            # Draw the cell borders
            pdf_canvas.rect(current_x, y_offset, column_widths[i], cell_height, stroke=1, fill=0)
            current_x += column_widths[i]
        y_offset -= cell_height

    # Finalize the PDF
    pdf_canvas.showPage()
    pdf_canvas.save()

    return response


def Historique(request):
    query = request.GET.get('q')
    bulletins = Bulletin.objects.filter(user=request.user)

    # Rechercher uniquement par numéro de bulletin
    if query:
        bulletins = bulletins.filter(id__icontains=query)


    return render(request, 'simulateur_paie/historique.html', {"bulletins": bulletins})

def Historique_net_to_brut(request):
    query = request.GET.get('q')
    net_to_brut = Net_to_Brut.objects.filter(user=request.user)

    # Rechercher uniquement par numéro de bulletin
    if query:
        net_to_brut = net_to_brut.filter(id__icontains=query)


    return render(request, 'simulateur_paie/historique_Net_to_Brut.html', {"net_to_brut": net_to_brut})

def Historique_cout_to_brut(request):
    query = request.GET.get('q')
    cout_to_brut = Cout_to_Brut.objects.filter(user=request.user)

    # Rechercher uniquement par numéro de bulletin
    if query:
        cout_to_brut = cout_to_brut.filter(id__icontains=query)


    return render(request, 'simulateur_paie/historique_Cout_to_Brut.html', {"cout_to_brut": cout_to_brut})



def bulletin_detail(request, bulletin_id):
    # Retrieve the specific bulletin
    bulletin = Bulletin.objects.get(pk = bulletin_id)

    excel_data_V2 = [[0 for _ in range(6)] for _ in range(38)]

    # Titres des colonnes
    excel_data_V2[0][0] = "Libellé"
    excel_data_V2[0][1] = "Base"
    excel_data_V2[0][2] = "Taux salarial"
    excel_data_V2[0][3] = "Montant salarial"
    excel_data_V2[0][4] = "Taux patronal"
    excel_data_V2[0][5] = "Montant patronal"

    # Titres des lignes
    excel_data_V2[2][0] = "SALAIRE DE BASE MENSUEL"
    excel_data_V2[3][0] = "VARIABLES DE REMUNERATION"

    excel_data_V2[5][0] = "TOTAL SALAIRE BRUT"

    excel_data_V2[7][0] = "CAR"
    excel_data_V2[8][0] = "CCSS"
    excel_data_V2[9][0] = "Réduction CCSS Gens de maison (33%)"
    excel_data_V2[10][0] = "POLE EMPLOI"
    excel_data_V2[11][0] = "RETRAITE COMPLEMENTAIRE CMRC T1"
    excel_data_V2[12][0] = "RETRAITE COMPLEMENTAIRE CMRC T2"
    excel_data_V2[13][0] = "PREVOYANCE MC TRANCHE A"
    excel_data_V2[14][0] = "PREVOYANCE MC TRANCHE B"
    excel_data_V2[15][0] = "PREVOYANCE FR TRANCHE A"
    excel_data_V2[16][0] = "PREVOYANCE FR TRANCHE B"
    excel_data_V2[17][0] = "ACCIDENT DU TRAVAIL"
    excel_data_V2[18][0] = "TAXE SUR COTISATION AT"
    excel_data_V2[19][0] = "CCPB OUVRIER"
    excel_data_V2[20][0] = "CCPB ETAM/CADRE"

    excel_data_V2[22][0] = "Total cotisations"

    excel_data_V2[24][0] = "Taux de charges"

    excel_data_V2[26][0] = "Non-Soumis"
    excel_data_V2[27][0] = "Indemnité 5% Monégasque"
    excel_data_V2[28][0] = "Remboursement Transport"
    excel_data_V2[29][0] = "Tickets Restaurant"
    excel_data_V2[30][0] = "Mutuelle"
    excel_data_V2[31][0] = "Indemnité Télétravail mensuelle"
    excel_data_V2[32][0] = "Paniers"

    excel_data_V2[34][0] = "Net à payer"

    excel_data_V2[36][0] = "COUT TOTAL MENSUEL EMPLOYEUR"
    excel_data_V2[37][0] = "COUT ANNUEL TOTAL EMPLOYEUR"


    # Taux et plafonds :

    CAR = 6028 # B8
    CAR_part_salariale = 0.0685 # C8
    CAR_part_patronale = 0.0831 # E8

    CCSS = 9600 # B9
    CCSS_part_patronale = 0.1345 # E9

    CCSS_taux_reduit = CCSS # B10
    CCSS_taux_reduit_part_patronale = 0.1340 # E10

    plafond_fracais_tranche_A = 3864 # B11

    SMIC = 11.65 # B12

    Horaire_temps_plein_monegasque = 169 # B13

    plafond_plein_CMRC = 151 # E13

    pole_emploi = plafond_fracais_tranche_A * 4 # B16
    pole_emploi_part_salariale = 0.024 # C16
    pole_emploi_part_patronale = 0.0405 # E16

    CMRC_tranche1 = 3947 # B17
    CMRC_tranche1_part_salariale = 0.0401 # C17
    CMRC_tranche1_part_patronale = 0.0601 # E17

    CMRC_tranche2 = CMRC_tranche1 * 7 # B18
    CMRC_tranche2_part_salariale = 0.0972 # C18
    CMRC_tranche2_part_patronale = 0.1457 # E18

    CCPB_OUVRIER_salariale = 0.004 # C19
    CCPB_OUVRIER_patronale = 0.201 # E19

    CCPB_ETAM_CADRE_salariale = 0.0 # C20
    CCPB_ETAM_CADRE_patronale = 0.205 # E20

    prevoyance_cadre_plafond_monégasque = CCSS # B21
    prevoyance_cadre_plafond_monégasque_patronale = 0.015 # E21

    prevoyance_cadre_plafond_francais =  plafond_fracais_tranche_A # B22
    prevoyance_cadre_plafond_francais_patronale = 0.015 # E22

    taxe_cotisation_AT_patronale = 0.04 # E23

    minimu_garanti = 4.15 # B25



    #ligne 31
    excel_data_V2[2][1] = round(bulletin.heures_mois,2) # B31
    excel_data_V2[2][2] = bulletin.salaire_brut_mois / excel_data_V2[2][1] # C31
    excel_data_V2[2][3] = bulletin.salaire_brut_mois # D31

    #ligne 34
    excel_data_V2[5][3] = excel_data_V2[2][3] # D34

    #ligne 36
    if excel_data_V2[5][3] < CAR:
        excel_data_V2[7][1] = excel_data_V2[5][3] # B36
    else:
        excel_data_V2[7][1] = CAR # B36

    excel_data_V2[7][2] = CAR_part_salariale # C36
    excel_data_V2[7][3] = excel_data_V2[7][1] * excel_data_V2[7][2] # D36
    excel_data_V2[7][4] = CAR_part_patronale # E36
    excel_data_V2[7][5] = excel_data_V2[7][1] * excel_data_V2[7][4] # F36

    #ligne 37
    if bulletin.admin_SAM == "Oui":
        excel_data_V2[8][1] = CCSS # B37
    else:
        excel_data_V2[8][1] = min(excel_data_V2[5][3], CCSS) # B37

    if bulletin.taux_ccss_red == "Oui":
        excel_data_V2[8][4] = CCSS_taux_reduit_part_patronale # E37
    else:
        excel_data_V2[8][4] = CCSS_part_patronale # E37

    excel_data_V2[8][5] = excel_data_V2[8][1] * excel_data_V2[8][4] # F37

    #ligne 38
    if bulletin.Gens_de_Maison == "Oui":
        excel_data_V2[9][1] = -excel_data_V2[8][1] * 0.67 # B38
    else:
        excel_data_V2[9][1] = 0 # B38

    excel_data_V2[9][4] = CCSS_taux_reduit_part_patronale # E38
    excel_data_V2[9][5] = excel_data_V2[9][1] * excel_data_V2[9][4] # F38

    #ligne 39
    H39 = (plafond_fracais_tranche_A / 169 * excel_data_V2[2][1]) * 4 #********************************************************

    if bulletin.Exclu_ass_cho == "Oui":
        excel_data_V2[10][1] = 0 # B39
    else:
        excel_data_V2[10][1] = min(excel_data_V2[5][3], H39) # B39

    excel_data_V2[10][2] = pole_emploi_part_salariale # C39
    excel_data_V2[10][3] = excel_data_V2[10][1] * excel_data_V2[10][2] # D39
    excel_data_V2[10][4] = pole_emploi_part_patronale # E39
    excel_data_V2[10][5] = excel_data_V2[10][1] * excel_data_V2[10][4] # F39

    #ligne 40
    rounded_B31 = round(excel_data_V2[2][1])  # Arrondi de B31 à 0 décimale
    calculation = (CMRC_tranche1 / plafond_plein_CMRC) * rounded_B31 
    H40 = min(CMRC_tranche1, calculation)

    if excel_data_V2[5][3] <= H40:
        excel_data_V2[11][1] = excel_data_V2[5][3] # B40
    else:
        excel_data_V2[11][1] = H40 # B40

    excel_data_V2[11][2] = CMRC_tranche1_part_salariale # C40
    excel_data_V2[11][3] = excel_data_V2[11][1] * excel_data_V2[11][2] # D40
    excel_data_V2[11][4] = CMRC_tranche1_part_patronale # E40
    excel_data_V2[11][5] = excel_data_V2[11][1] * excel_data_V2[11][4] # F40

    #ligne 41
    H41 = H40*7

    if excel_data_V2[5][3] <= (H40 + H41):
        excel_data_V2[12][1] = excel_data_V2[5][3] - excel_data_V2[11][1] # B41
    else:
        excel_data_V2[12][1] = (H40 + H41) - excel_data_V2[11][1] # B41
    
    excel_data_V2[12][2] = CMRC_tranche2_part_salariale  # C41
    excel_data_V2[12][3] = excel_data_V2[12][1] * excel_data_V2[12][2] # D41
    excel_data_V2[12][4] = CMRC_tranche2_part_patronale # E41
    excel_data_V2[12][5] = excel_data_V2[12][1] * excel_data_V2[12][4] # F41

    #ligne 42
    B8 = "Oui" if bulletin.prevoyance else "Non"
    H42 = prevoyance_cadre_plafond_monégasque / Horaire_temps_plein_monegasque * excel_data_V2[2][1]

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[13][1] = 0 # B42
        else:
            excel_data_V2[13][1] = min(H42, excel_data_V2[5][3]) # B42
    else:
        excel_data_V2[13][1] = 0 # B42
    
    excel_data_V2[13][2] = bulletin.taux_prev_T1_sal / 100 # C42
    excel_data_V2[13][3] = excel_data_V2[13][1] * excel_data_V2[13][2]  # D42
    excel_data_V2[13][4] = bulletin.taux_prev_T1_patr / 100 # E42
    excel_data_V2[13][5] = excel_data_V2[13][1] * excel_data_V2[13][4] # F42

    #ligne 43
    H43 = H39 - H42

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[14][1] = 0 # B43
        else:
            if excel_data_V2[5][3] > H42:
                excel_data_V2[14][1] = min(excel_data_V2[5][3]-excel_data_V2[13][1],H42+H43-excel_data_V2[13][1]) # B43
            else:
                excel_data_V2[14][1] = 0 # B43
    else:
        excel_data_V2[14][1] = 0 # B43
    
    excel_data_V2[14][2] =  bulletin.taux_prev_T2_sal / 100 # C43
    excel_data_V2[14][3] =  excel_data_V2[14][1] * excel_data_V2[14][2] # D43
    excel_data_V2[14][4] =  bulletin.taux_prev_T2_patr / 100 # E43
    excel_data_V2[14][5] =  excel_data_V2[14][1] * excel_data_V2[14][4] # F43

    #ligne 44
    H46 = plafond_fracais_tranche_A / Horaire_temps_plein_monegasque * excel_data_V2[2][1]

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[15][1] = min(excel_data_V2[5][3],H46) # B44
        else:
            excel_data_V2[15][1] = 0 # B44
    else:
        excel_data_V2[15][1] = 0 # B44

    excel_data_V2[15][2] = bulletin.taux_prev_T1_sal / 100 # C44
    excel_data_V2[15][3] = excel_data_V2[15][1] * excel_data_V2[15][2]  # D44
    excel_data_V2[15][4] = bulletin.taux_prev_T1_patr / 100 # E44
    excel_data_V2[15][5] = excel_data_V2[15][1] * excel_data_V2[15][4] # F44

    #ligne 45
    H49 = H46*4

    if B8 == "Non":
        excel_data_V2[16][1] = 0 # B45
    else:
        if bulletin.type_plafond != "FR URSSAF":
            excel_data_V2[16][1]  = 0 # B45
        else:
            if excel_data_V2[5][3] > H46: 
                excel_data_V2[16][1] = min(excel_data_V2[5][3] - excel_data_V2[15][1], H49 - excel_data_V2[15][1]) # B45
            else:
                excel_data_V2[16][1] = 0 # B45

    excel_data_V2[16][2] = bulletin.taux_prev_T2_sal / 100 # C45
    excel_data_V2[16][3] = excel_data_V2[16][1] * excel_data_V2[16][2] # D45
    excel_data_V2[16][4] = bulletin.taux_prev_T2_patr / 100 # E45
    excel_data_V2[16][3] = excel_data_V2[16][1] * excel_data_V2[16][4] # F45


    #ligne 46
    excel_data_V2[17][1] =  excel_data_V2[5][3] # B46

    if bulletin.Taux_ATHT > 0:
        excel_data_V2[17][4] = bulletin.Taux_ATHT / 100 # E46
    else:
        excel_data_V2[17][4] = 1 / 100 # E46

    excel_data_V2[17][5] = excel_data_V2[17][1] * excel_data_V2[17][4] # F46

    #ligne 47
    excel_data_V2[18][1] = excel_data_V2[17][5] # B47

    excel_data_V2[18][4] = taxe_cotisation_AT_patronale # E47

    excel_data_V2[18][5] = excel_data_V2[18][1] * excel_data_V2[18][4] # F47

    #ligne 48
    if bulletin.CCPB_ouvrier == "Oui":
        excel_data_V2[19][1] = excel_data_V2[5][3] # B48
    else:
        excel_data_V2[19][1] = 0 # B48

    excel_data_V2[19][2] = CCPB_OUVRIER_salariale # C48

    excel_data_V2[19][3] = excel_data_V2[19][1] * excel_data_V2[19][2] # D48
    excel_data_V2[19][4] = CCPB_OUVRIER_patronale # E48
    excel_data_V2[19][5] = excel_data_V2[19][1] * excel_data_V2[19][4] # F48
    

    #ligne 49

    if bulletin.CCPB_Etam_cadre == "Oui":
        excel_data_V2[20][1] = excel_data_V2[5][3] # B49
    else:
        excel_data_V2[20][1] = 0 # B49

    excel_data_V2[20][2] = CCPB_ETAM_CADRE_salariale # C49 

    excel_data_V2[20][3] = excel_data_V2[20][1] * excel_data_V2[20][2] # D49
    excel_data_V2[20][4] = CCPB_ETAM_CADRE_patronale # E49
    excel_data_V2[20][5] = excel_data_V2[20][1] * excel_data_V2[20][4] # F49


    #ligne 51
    excel_data_V2[22][3] = 0
    for i in range(7,21):
        excel_data_V2[22][3] = excel_data_V2[22][3] + excel_data_V2[i][3] # D51

    excel_data_V2[22][5] = 0
    for i in range(7,21):
        excel_data_V2[22][5] = excel_data_V2[22][5] + excel_data_V2[i][5] # F51

    #ligne 53
    excel_data_V2[24][2] = excel_data_V2[22][3] / excel_data_V2[5][3] # C53

    excel_data_V2[24][4] = excel_data_V2[22][5] / excel_data_V2[5][3] # E53

    #ligne 56
    H50 = SMIC * 1.05

    if excel_data_V2[2][2] <= H50:
        excel_data_V2[27][1] = excel_data_V2[5][3] # B56
    else:
        excel_data_V2[27][1] = 0 # B56

    excel_data_V2[27][2] = 0.05 # C56

    excel_data_V2[27][3] = excel_data_V2[27][1] * excel_data_V2[27][2] # D56

    excel_data_V2[27][5] = excel_data_V2[27][3] # F56

    #ligne 57
    excel_data_V2[28][3] = bulletin.remboursement_transport # D57

    excel_data_V2[28][5] = excel_data_V2[28][3] # F57

    #ligne 58
    excel_data_V2[29][1] = bulletin.ticket_resto_prix # B58
    excel_data_V2[29][2] = -bulletin.ticket_resto_sal # C58
    excel_data_V2[29][3] = excel_data_V2[29][1] * excel_data_V2[29][2] # D58
    excel_data_V2[29][4] = bulletin.ticket_resto_patr # E58
    excel_data_V2[29][5] = excel_data_V2[29][4] * excel_data_V2[29][1] # F58

    #ligne 59
    B12 = "Oui" if bulletin.mutuelle else "Non"

    if B12 == "Oui":
        excel_data_V2[30][3] = -bulletin.montant_mutu_sal # D59
    else:
        excel_data_V2[30][3] = 0 # D59

    if B12 == "Oui":
        excel_data_V2[30][5] = bulletin.montant_mutu_patr # F59
    else:
        excel_data_V2[30][5] = 0 # F59

    #ligne 60
    excel_data_V2[31][3] = 20 # D60

    #ligne 61
    excel_data_V2[32][1] = bulletin.panier_prix # B61
    excel_data_V2[32][2] = bulletin.paniers_sal # C61
    excel_data_V2[32][3] = excel_data_V2[32][1] * excel_data_V2[32][2] # D61
    excel_data_V2[32][5] = excel_data_V2[32][3] # F61

    #ligne 62
    excel_data_V2[33][5] = 0
    for i in range(22,33):
        excel_data_V2[33][5] = excel_data_V2[33][5] + excel_data_V2[i][5] # F62

    #ligne 63
    somme = 0
    for i in range(24,33):
        somme = somme + excel_data_V2[i][3]

    excel_data_V2[34][3] = excel_data_V2[2][3] - excel_data_V2[22][3] + somme # D63

    #ligne 65
    excel_data_V2[36][1] = excel_data_V2[33][5] + excel_data_V2[2][3] # B65

    #ligne 66
    excel_data_V2[37][1] = excel_data_V2[36][1]*12 # B66


    print(len(excel_data_V2))
    print(len(excel_data_V2[0]))
    print(type(excel_data_V2))

    # Formatage des datas (arrondis, %, ...)
    for i in range(7, 28):
        if excel_data_V2[i][2] is not None:
            excel_data_V2[i][2] = excel_data_V2[i][2] * 100

    for i in range(7, 25):
        if excel_data_V2[i][4] is not None:
            excel_data_V2[i][4] = excel_data_V2[i][4] * 100

    for i in range(2,38):
        for j in range(1,6):
            if (excel_data_V2[i][j] is not None) and not isinstance(excel_data_V2[i][j], str):
                if (i== 2) and (j==2):
                    excel_data_V2[i][j] = f"{excel_data_V2[i][j]:.4f}" # attention ça fait que maintenant les chiffres deviennent des strings
                else:
                    excel_data_V2[i][j] = f"{excel_data_V2[i][j]:.2f}" # attention ça fait que maintenant les chiffres deviennent des strings

                if (excel_data_V2[i][j] == f"{0.0:.2f}"):
                    excel_data_V2[i][j] = " "

                excel_data_V2[i][j] = str(excel_data_V2[i][j])

                for p in range(0, len(excel_data_V2[i][j])):
                    if (excel_data_V2[i][j][p] == ".") and (len(excel_data_V2[i][j][:p]) >= 4) and (len(excel_data_V2[i][j][:p]) < 7):
                        u = p - 3
                        excel_data_V2[i][j] = excel_data_V2[i][j][:u] + " " + excel_data_V2[i][j][u:]
                                

    excel_data_V2[34][3] = excel_data_V2[34][3] + " €"
    excel_data_V2[36][1] = excel_data_V2[36][1] + " €"
    excel_data_V2[37][1] = excel_data_V2[37][1] + " €"

    for i in range(7, 28):
        if excel_data_V2[i][2] is not None and excel_data_V2[i][2] != " ":
            excel_data_V2[i][2] = str(excel_data_V2[i][2]) + "%"

    for i in range(7, 25):
        if excel_data_V2[i][4] is not None and excel_data_V2[i][4] != " ":
            excel_data_V2[i][4] = str(excel_data_V2[i][4]) + "%"

    

    prev_bulletin = Bulletin.objects.filter(id__lt=bulletin_id).order_by('-id').first()
    next_bulletin = Bulletin.objects.filter(id__gt=bulletin_id).order_by('id').first()


    # Pass the formatted data to the template
    context = {'excel_data': excel_data_V2,
        "bulletin_ID" : bulletin_id,
        'bulletin': bulletin,
        'prev_bulletin': prev_bulletin,
        'next_bulletin': next_bulletin
    }

    return render(request, 'simulateur_paie/bulletin.html', context)

def bulletin_detail_net_to_brut(request, bulletin_id):
    # Retrieve the specific bulletin

    net_vers_brut = Net_to_Brut.objects.get(pk = bulletin_id)

    bulletin = net_vers_brut.bulletin

    excel_data_V2 = [[0 for _ in range(6)] for _ in range(38)]

    # Titres des colonnes
    excel_data_V2[0][0] = "Libellé"
    excel_data_V2[0][1] = "Base"
    excel_data_V2[0][2] = "Taux salarial"
    excel_data_V2[0][3] = "Montant salarial"
    excel_data_V2[0][4] = "Taux patronal"
    excel_data_V2[0][5] = "Montant patronal"

    # Titres des lignes
    excel_data_V2[2][0] = "SALAIRE DE BASE MENSUEL"
    excel_data_V2[3][0] = "VARIABLES DE REMUNERATION"

    excel_data_V2[5][0] = "TOTAL SALAIRE BRUT"

    excel_data_V2[7][0] = "CAR"
    excel_data_V2[8][0] = "CCSS"
    excel_data_V2[9][0] = "Réduction CCSS Gens de maison (33%)"
    excel_data_V2[10][0] = "POLE EMPLOI"
    excel_data_V2[11][0] = "RETRAITE COMPLEMENTAIRE CMRC T1"
    excel_data_V2[12][0] = "RETRAITE COMPLEMENTAIRE CMRC T2"
    excel_data_V2[13][0] = "PREVOYANCE MC TRANCHE A"
    excel_data_V2[14][0] = "PREVOYANCE MC TRANCHE B"
    excel_data_V2[15][0] = "PREVOYANCE FR TRANCHE A"
    excel_data_V2[16][0] = "PREVOYANCE FR TRANCHE B"
    excel_data_V2[17][0] = "ACCIDENT DU TRAVAIL"
    excel_data_V2[18][0] = "TAXE SUR COTISATION AT"
    excel_data_V2[19][0] = "CCPB OUVRIER"
    excel_data_V2[20][0] = "CCPB ETAM/CADRE"

    excel_data_V2[22][0] = "Total cotisations"

    excel_data_V2[24][0] = "Taux de charges"

    excel_data_V2[26][0] = "Non-Soumis"
    excel_data_V2[27][0] = "Indemnité 5% Monégasque"
    excel_data_V2[28][0] = "Remboursement Transport"
    excel_data_V2[29][0] = "Tickets Restaurant"
    excel_data_V2[30][0] = "Mutuelle"
    excel_data_V2[31][0] = "Indemnité Télétravail mensuelle"
    excel_data_V2[32][0] = "Paniers"

    excel_data_V2[34][0] = "Net à payer"

    excel_data_V2[36][0] = "COUT TOTAL MENSUEL EMPLOYEUR"
    excel_data_V2[37][0] = "COUT ANNUEL TOTAL EMPLOYEUR"


    # Taux et plafonds :

    CAR = 6028 # B8
    CAR_part_salariale = 0.0685 # C8
    CAR_part_patronale = 0.0831 # E8

    CCSS = 9600 # B9
    CCSS_part_patronale = 0.1345 # E9

    CCSS_taux_reduit = CCSS # B10
    CCSS_taux_reduit_part_patronale = 0.1340 # E10

    plafond_fracais_tranche_A = 3864 # B11

    SMIC = 11.65 # B12

    Horaire_temps_plein_monegasque = 169 # B13

    plafond_plein_CMRC = 151 # E13

    pole_emploi = plafond_fracais_tranche_A * 4 # B16
    pole_emploi_part_salariale = 0.024 # C16
    pole_emploi_part_patronale = 0.0405 # E16

    CMRC_tranche1 = 3947 # B17
    CMRC_tranche1_part_salariale = 0.0401 # C17
    CMRC_tranche1_part_patronale = 0.0601 # E17

    CMRC_tranche2 = CMRC_tranche1 * 7 # B18
    CMRC_tranche2_part_salariale = 0.0972 # C18
    CMRC_tranche2_part_patronale = 0.1457 # E18

    CCPB_OUVRIER_salariale = 0.004 # C19
    CCPB_OUVRIER_patronale = 0.201 # E19

    CCPB_ETAM_CADRE_salariale = 0.0 # C20
    CCPB_ETAM_CADRE_patronale = 0.205 # E20

    prevoyance_cadre_plafond_monégasque = CCSS # B21
    prevoyance_cadre_plafond_monégasque_patronale = 0.015 # E21

    prevoyance_cadre_plafond_francais =  plafond_fracais_tranche_A # B22
    prevoyance_cadre_plafond_francais_patronale = 0.015 # E22

    taxe_cotisation_AT_patronale = 0.04 # E23

    minimu_garanti = 4.15 # B25



    #ligne 31
    excel_data_V2[2][1] = round(bulletin.heures_mois,2) # B31
    excel_data_V2[2][2] = net_vers_brut.salaire_brut / excel_data_V2[2][1] # C31
    excel_data_V2[2][3] = net_vers_brut.salaire_brut # D31

    #ligne 34
    excel_data_V2[5][3] = excel_data_V2[2][3] # D34

    #ligne 36
    if excel_data_V2[5][3] < CAR:
        excel_data_V2[7][1] = excel_data_V2[5][3] # B36
    else:
        excel_data_V2[7][1] = CAR # B36

    excel_data_V2[7][2] = CAR_part_salariale # C36
    excel_data_V2[7][3] = excel_data_V2[7][1] * excel_data_V2[7][2] # D36
    excel_data_V2[7][4] = CAR_part_patronale # E36
    excel_data_V2[7][5] = excel_data_V2[7][1] * excel_data_V2[7][4] # F36

    #ligne 37
    if bulletin.admin_SAM == "Oui":
        excel_data_V2[8][1] = CCSS # B37
    else:
        excel_data_V2[8][1] = min(excel_data_V2[5][3], CCSS) # B37

    if bulletin.taux_ccss_red == "Oui":
        excel_data_V2[8][4] = CCSS_taux_reduit_part_patronale # E37
    else:
        excel_data_V2[8][4] = CCSS_part_patronale # E37

    excel_data_V2[8][5] = excel_data_V2[8][1] * excel_data_V2[8][4] # F37

    #ligne 38
    if bulletin.Gens_de_Maison == "Oui":
        excel_data_V2[9][1] = -excel_data_V2[8][1] * 0.67 # B38
    else:
        excel_data_V2[9][1] = 0 # B38

    excel_data_V2[9][4] = CCSS_taux_reduit_part_patronale # E38
    excel_data_V2[9][5] = excel_data_V2[9][1] * excel_data_V2[9][4] # F38

    #ligne 39
    H39 = (plafond_fracais_tranche_A / 169 * excel_data_V2[2][1]) * 4 #********************************************************

    if bulletin.Exclu_ass_cho == "Oui":
        excel_data_V2[10][1] = 0 # B39
    else:
        excel_data_V2[10][1] = min(excel_data_V2[5][3], H39) # B39

    excel_data_V2[10][2] = pole_emploi_part_salariale # C39
    excel_data_V2[10][3] = excel_data_V2[10][1] * excel_data_V2[10][2] # D39
    excel_data_V2[10][4] = pole_emploi_part_patronale # E39
    excel_data_V2[10][5] = excel_data_V2[10][1] * excel_data_V2[10][4] # F39

    #ligne 40
    rounded_B31 = round(excel_data_V2[2][1])  # Arrondi de B31 à 0 décimale
    calculation = (CMRC_tranche1 / plafond_plein_CMRC) * rounded_B31 
    H40 = min(CMRC_tranche1, calculation)

    if excel_data_V2[5][3] <= H40:
        excel_data_V2[11][1] = excel_data_V2[5][3] # B40
    else:
        excel_data_V2[11][1] = H40 # B40

    excel_data_V2[11][2] = CMRC_tranche1_part_salariale # C40
    excel_data_V2[11][3] = excel_data_V2[11][1] * excel_data_V2[11][2] # D40
    excel_data_V2[11][4] = CMRC_tranche1_part_patronale # E40
    excel_data_V2[11][5] = excel_data_V2[11][1] * excel_data_V2[11][4] # F40

    #ligne 41
    H41 = H40*7

    if excel_data_V2[5][3] <= (H40 + H41):
        excel_data_V2[12][1] = excel_data_V2[5][3] - excel_data_V2[11][1] # B41
    else:
        excel_data_V2[12][1] = (H40 + H41) - excel_data_V2[11][1] # B41
    
    excel_data_V2[12][2] = CMRC_tranche2_part_salariale  # C41
    excel_data_V2[12][3] = excel_data_V2[12][1] * excel_data_V2[12][2] # D41
    excel_data_V2[12][4] = CMRC_tranche2_part_patronale # E41
    excel_data_V2[12][5] = excel_data_V2[12][1] * excel_data_V2[12][4] # F41

    #ligne 42
    B8 = "Oui" if bulletin.prevoyance else "Non"
    H42 = prevoyance_cadre_plafond_monégasque / Horaire_temps_plein_monegasque * excel_data_V2[2][1]

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[13][1] = 0 # B42
        else:
            excel_data_V2[13][1] = min(H42, excel_data_V2[5][3]) # B42
    else:
        excel_data_V2[13][1] = 0 # B42
    
    excel_data_V2[13][2] = bulletin.taux_prev_T1_sal / 100 # C42
    excel_data_V2[13][3] = excel_data_V2[13][1] * excel_data_V2[13][2]  # D42
    excel_data_V2[13][4] = bulletin.taux_prev_T1_patr / 100 # E42
    excel_data_V2[13][5] = excel_data_V2[13][1] * excel_data_V2[13][4] # F42

    #ligne 43
    H43 = H39 - H42

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[14][1] = 0 # B43
        else:
            if excel_data_V2[5][3] > H42:
                excel_data_V2[14][1] = min(excel_data_V2[5][3]-excel_data_V2[13][1],H42+H43-excel_data_V2[13][1]) # B43
            else:
                excel_data_V2[14][1] = 0 # B43
    else:
        excel_data_V2[14][1] = 0 # B43
    
    excel_data_V2[14][2] =  bulletin.taux_prev_T2_sal / 100 # C43
    excel_data_V2[14][3] =  excel_data_V2[14][1] * excel_data_V2[14][2] # D43
    excel_data_V2[14][4] =  bulletin.taux_prev_T2_patr / 100 # E43
    excel_data_V2[14][5] =  excel_data_V2[14][1] * excel_data_V2[14][4] # F43

    #ligne 44
    H46 = plafond_fracais_tranche_A / Horaire_temps_plein_monegasque * excel_data_V2[2][1]

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[15][1] = min(excel_data_V2[5][3],H46) # B44
        else:
            excel_data_V2[15][1] = 0 # B44
    else:
        excel_data_V2[15][1] = 0 # B44

    excel_data_V2[15][2] = bulletin.taux_prev_T1_sal / 100 # C44
    excel_data_V2[15][3] = excel_data_V2[15][1] * excel_data_V2[15][2]  # D44
    excel_data_V2[15][4] = bulletin.taux_prev_T1_patr / 100 # E44
    excel_data_V2[15][5] = excel_data_V2[15][1] * excel_data_V2[15][4] # F44

    #ligne 45
    H49 = H46*4

    if B8 == "Non":
        excel_data_V2[16][1] = 0 # B45
    else:
        if bulletin.type_plafond != "FR URSSAF":
            excel_data_V2[16][1]  = 0 # B45
        else:
            if excel_data_V2[5][3] > H46: 
                excel_data_V2[16][1] = min(excel_data_V2[5][3] - excel_data_V2[15][1], H49 - excel_data_V2[15][1]) # B45
            else:
                excel_data_V2[16][1] = 0 # B45

    excel_data_V2[16][2] = bulletin.taux_prev_T2_sal / 100 # C45
    excel_data_V2[16][3] = excel_data_V2[16][1] * excel_data_V2[16][2] # D45
    excel_data_V2[16][4] = bulletin.taux_prev_T2_patr / 100 # E45
    excel_data_V2[16][3] = excel_data_V2[16][1] * excel_data_V2[16][4] # F45


    #ligne 46
    excel_data_V2[17][1] =  excel_data_V2[5][3] # B46

    if bulletin.Taux_ATHT > 0:
        excel_data_V2[17][4] = bulletin.Taux_ATHT / 100 # E46
    else:
        excel_data_V2[17][4] = 1 / 100 # E46

    excel_data_V2[17][5] = excel_data_V2[17][1] * excel_data_V2[17][4] # F46

    #ligne 47
    excel_data_V2[18][1] = excel_data_V2[17][5] # B47

    excel_data_V2[18][4] = taxe_cotisation_AT_patronale # E47

    excel_data_V2[18][5] = excel_data_V2[18][1] * excel_data_V2[18][4] # F47

    #ligne 48
    if bulletin.CCPB_ouvrier == "Oui":
        excel_data_V2[19][1] = excel_data_V2[5][3] # B48
    else:
        excel_data_V2[19][1] = 0 # B48

    excel_data_V2[19][2] = CCPB_OUVRIER_salariale # C48

    excel_data_V2[19][3] = excel_data_V2[19][1] * excel_data_V2[19][2] # D48
    excel_data_V2[19][4] = CCPB_OUVRIER_patronale # E48
    excel_data_V2[19][5] = excel_data_V2[19][1] * excel_data_V2[19][4] # F48
    

    #ligne 49

    if bulletin.CCPB_Etam_cadre == "Oui":
        excel_data_V2[20][1] = excel_data_V2[5][3] # B49
    else:
        excel_data_V2[20][1] = 0 # B49

    excel_data_V2[20][2] = CCPB_ETAM_CADRE_salariale # C49 

    excel_data_V2[20][3] = excel_data_V2[20][1] * excel_data_V2[20][2] # D49
    excel_data_V2[20][4] = CCPB_ETAM_CADRE_patronale # E49
    excel_data_V2[20][5] = excel_data_V2[20][1] * excel_data_V2[20][4] # F49


    #ligne 51
    excel_data_V2[22][3] = 0
    for i in range(7,21):
        excel_data_V2[22][3] = excel_data_V2[22][3] + excel_data_V2[i][3] # D51

    excel_data_V2[22][5] = 0
    for i in range(7,21):
        excel_data_V2[22][5] = excel_data_V2[22][5] + excel_data_V2[i][5] # F51

    #ligne 53
    excel_data_V2[24][2] = excel_data_V2[22][3] / excel_data_V2[5][3] # C53

    excel_data_V2[24][4] = excel_data_V2[22][5] / excel_data_V2[5][3] # E53

    #ligne 56
    H50 = SMIC * 1.05

    if excel_data_V2[2][2] <= H50:
        excel_data_V2[27][1] = excel_data_V2[5][3] # B56
    else:
        excel_data_V2[27][1] = 0 # B56

    excel_data_V2[27][2] = 0.05 # C56

    excel_data_V2[27][3] = excel_data_V2[27][1] * excel_data_V2[27][2] # D56

    excel_data_V2[27][5] = excel_data_V2[27][3] # F56

    #ligne 57
    excel_data_V2[28][3] = bulletin.remboursement_transport # D57

    excel_data_V2[28][5] = excel_data_V2[28][3] # F57

    #ligne 58
    excel_data_V2[29][1] = bulletin.ticket_resto_prix # B58
    excel_data_V2[29][2] = -bulletin.ticket_resto_sal # C58
    excel_data_V2[29][3] = excel_data_V2[29][1] * excel_data_V2[29][2] # D58
    excel_data_V2[29][4] = bulletin.ticket_resto_patr # E58
    excel_data_V2[29][5] = excel_data_V2[29][4] * excel_data_V2[29][1] # F58

    #ligne 59
    B12 = "Oui" if bulletin.mutuelle else "Non"

    if B12 == "Oui":
        excel_data_V2[30][3] = -bulletin.montant_mutu_sal # D59
    else:
        excel_data_V2[30][3] = 0 # D59

    if B12 == "Oui":
        excel_data_V2[30][5] = bulletin.montant_mutu_patr # F59
    else:
        excel_data_V2[30][5] = 0 # F59

    #ligne 60
    excel_data_V2[31][3] = 20 # D60

    #ligne 61
    excel_data_V2[32][1] = bulletin.panier_prix # B61
    excel_data_V2[32][2] = bulletin.paniers_sal # C61
    excel_data_V2[32][3] = excel_data_V2[32][1] * excel_data_V2[32][2] # D61
    excel_data_V2[32][5] = excel_data_V2[32][3] # F61

    #ligne 62
    excel_data_V2[33][5] = 0
    for i in range(22,33):
        excel_data_V2[33][5] = excel_data_V2[33][5] + excel_data_V2[i][5] # F62

    #ligne 63
    somme = 0
    for i in range(24,33):
        somme = somme + excel_data_V2[i][3]

    excel_data_V2[34][3] = excel_data_V2[2][3] - excel_data_V2[22][3] + somme # D63

    #ligne 65
    excel_data_V2[36][1] = excel_data_V2[33][5] + excel_data_V2[2][3] # B65

    #ligne 66
    excel_data_V2[37][1] = excel_data_V2[36][1]*12 # B66


    print(len(excel_data_V2))
    print(len(excel_data_V2[0]))
    print(type(excel_data_V2))

    # Formatage des datas (arrondis, %, ...)
    for i in range(7, 28):
        if excel_data_V2[i][2] is not None:
            excel_data_V2[i][2] = excel_data_V2[i][2] * 100

    for i in range(7, 25):
        if excel_data_V2[i][4] is not None:
            excel_data_V2[i][4] = excel_data_V2[i][4] * 100

    for i in range(2,38):
        for j in range(1,6):
            if (excel_data_V2[i][j] is not None) and not isinstance(excel_data_V2[i][j], str):
                if (i== 2) and (j==2):
                    excel_data_V2[i][j] = f"{excel_data_V2[i][j]:.4f}" # attention ça fait que maintenant les chiffres deviennent des strings
                else:
                    excel_data_V2[i][j] = f"{excel_data_V2[i][j]:.2f}" # attention ça fait que maintenant les chiffres deviennent des strings

                if (excel_data_V2[i][j] == f"{0.0:.2f}"):
                    excel_data_V2[i][j] = " "

                excel_data_V2[i][j] = str(excel_data_V2[i][j])

                for p in range(0, len(excel_data_V2[i][j])):
                    if (excel_data_V2[i][j][p] == ".") and (len(excel_data_V2[i][j][:p]) >= 4) and (len(excel_data_V2[i][j][:p]) < 7):
                        u = p - 3
                        excel_data_V2[i][j] = excel_data_V2[i][j][:u] + " " + excel_data_V2[i][j][u:]
                                

    excel_data_V2[34][3] = excel_data_V2[34][3] + " €"
    excel_data_V2[36][1] = excel_data_V2[36][1] + " €"
    excel_data_V2[37][1] = excel_data_V2[37][1] + " €"

    for i in range(7, 28):
        if excel_data_V2[i][2] is not None and excel_data_V2[i][2] != " ":
            excel_data_V2[i][2] = str(excel_data_V2[i][2]) + "%"

    for i in range(7, 25):
        if excel_data_V2[i][4] is not None and excel_data_V2[i][4] != " ":
            excel_data_V2[i][4] = str(excel_data_V2[i][4]) + "%"

    

    prev_bulletin = Net_to_Brut.objects.filter(id__lt=bulletin_id).order_by('-id').first()
    next_bulletin = Net_to_Brut.objects.filter(id__gt=bulletin_id).order_by('id').first()

    
    brut_heure = net_vers_brut.salaire_brut / bulletin.heures_mois 

    brut_heure = round(brut_heure,4)


    # Pass the formatted data to the template
    context = {'excel_data': excel_data_V2,
        "bulletin_ID" : bulletin_id,
        'bulletin': bulletin,
        'prev_bulletin': prev_bulletin,
        'next_bulletin': next_bulletin,
        "net_to_brut": net_vers_brut,
        "brut_heure": brut_heure
    }

    return render(request, 'simulateur_paie/rendu_Net_to_brut.html', context)

def bulletin_detail_cout_to_brut(request, bulletin_id):
    # Retrieve the specific bulletin

    cout_vers_brut = Cout_to_Brut.objects.get(pk = bulletin_id)

    bulletin = cout_vers_brut.bulletin

    excel_data_V2 = [[0 for _ in range(6)] for _ in range(38)]

    # Titres des colonnes
    excel_data_V2[0][0] = "Libellé"
    excel_data_V2[0][1] = "Base"
    excel_data_V2[0][2] = "Taux salarial"
    excel_data_V2[0][3] = "Montant salarial"
    excel_data_V2[0][4] = "Taux patronal"
    excel_data_V2[0][5] = "Montant patronal"

    # Titres des lignes
    excel_data_V2[2][0] = "SALAIRE DE BASE MENSUEL"
    excel_data_V2[3][0] = "VARIABLES DE REMUNERATION"

    excel_data_V2[5][0] = "TOTAL SALAIRE BRUT"

    excel_data_V2[7][0] = "CAR"
    excel_data_V2[8][0] = "CCSS"
    excel_data_V2[9][0] = "Réduction CCSS Gens de maison (33%)"
    excel_data_V2[10][0] = "POLE EMPLOI"
    excel_data_V2[11][0] = "RETRAITE COMPLEMENTAIRE CMRC T1"
    excel_data_V2[12][0] = "RETRAITE COMPLEMENTAIRE CMRC T2"
    excel_data_V2[13][0] = "PREVOYANCE MC TRANCHE A"
    excel_data_V2[14][0] = "PREVOYANCE MC TRANCHE B"
    excel_data_V2[15][0] = "PREVOYANCE FR TRANCHE A"
    excel_data_V2[16][0] = "PREVOYANCE FR TRANCHE B"
    excel_data_V2[17][0] = "ACCIDENT DU TRAVAIL"
    excel_data_V2[18][0] = "TAXE SUR COTISATION AT"
    excel_data_V2[19][0] = "CCPB OUVRIER"
    excel_data_V2[20][0] = "CCPB ETAM/CADRE"

    excel_data_V2[22][0] = "Total cotisations"

    excel_data_V2[24][0] = "Taux de charges"

    excel_data_V2[26][0] = "Non-Soumis"
    excel_data_V2[27][0] = "Indemnité 5% Monégasque"
    excel_data_V2[28][0] = "Remboursement Transport"
    excel_data_V2[29][0] = "Tickets Restaurant"
    excel_data_V2[30][0] = "Mutuelle"
    excel_data_V2[31][0] = "Indemnité Télétravail mensuelle"
    excel_data_V2[32][0] = "Paniers"

    excel_data_V2[34][0] = "Net à payer"

    excel_data_V2[36][0] = "COUT TOTAL MENSUEL EMPLOYEUR"
    excel_data_V2[37][0] = "COUT ANNUEL TOTAL EMPLOYEUR"


    # Taux et plafonds :

    CAR = 6028 # B8
    CAR_part_salariale = 0.0685 # C8
    CAR_part_patronale = 0.0831 # E8

    CCSS = 9600 # B9
    CCSS_part_patronale = 0.1345 # E9

    CCSS_taux_reduit = CCSS # B10
    CCSS_taux_reduit_part_patronale = 0.1340 # E10

    plafond_fracais_tranche_A = 3864 # B11

    SMIC = 11.65 # B12

    Horaire_temps_plein_monegasque = 169 # B13

    plafond_plein_CMRC = 151 # E13

    pole_emploi = plafond_fracais_tranche_A * 4 # B16
    pole_emploi_part_salariale = 0.024 # C16
    pole_emploi_part_patronale = 0.0405 # E16

    CMRC_tranche1 = 3947 # B17
    CMRC_tranche1_part_salariale = 0.0401 # C17
    CMRC_tranche1_part_patronale = 0.0601 # E17

    CMRC_tranche2 = CMRC_tranche1 * 7 # B18
    CMRC_tranche2_part_salariale = 0.0972 # C18
    CMRC_tranche2_part_patronale = 0.1457 # E18

    CCPB_OUVRIER_salariale = 0.004 # C19
    CCPB_OUVRIER_patronale = 0.201 # E19

    CCPB_ETAM_CADRE_salariale = 0.0 # C20
    CCPB_ETAM_CADRE_patronale = 0.205 # E20

    prevoyance_cadre_plafond_monégasque = CCSS # B21
    prevoyance_cadre_plafond_monégasque_patronale = 0.015 # E21

    prevoyance_cadre_plafond_francais =  plafond_fracais_tranche_A # B22
    prevoyance_cadre_plafond_francais_patronale = 0.015 # E22

    taxe_cotisation_AT_patronale = 0.04 # E23

    minimu_garanti = 4.15 # B25



    #ligne 31
    excel_data_V2[2][1] = round(bulletin.heures_mois,2) # B31
    excel_data_V2[2][2] = cout_vers_brut.salaire_brut / excel_data_V2[2][1] # C31
    excel_data_V2[2][3] = cout_vers_brut.salaire_brut # D31

    #ligne 34
    excel_data_V2[5][3] = excel_data_V2[2][3] # D34

    #ligne 36
    if excel_data_V2[5][3] < CAR:
        excel_data_V2[7][1] = excel_data_V2[5][3] # B36
    else:
        excel_data_V2[7][1] = CAR # B36

    excel_data_V2[7][2] = CAR_part_salariale # C36
    excel_data_V2[7][3] = excel_data_V2[7][1] * excel_data_V2[7][2] # D36
    excel_data_V2[7][4] = CAR_part_patronale # E36
    excel_data_V2[7][5] = excel_data_V2[7][1] * excel_data_V2[7][4] # F36

    #ligne 37
    if bulletin.admin_SAM == "Oui":
        excel_data_V2[8][1] = CCSS # B37
    else:
        excel_data_V2[8][1] = min(excel_data_V2[5][3], CCSS) # B37

    if bulletin.taux_ccss_red == "Oui":
        excel_data_V2[8][4] = CCSS_taux_reduit_part_patronale # E37
    else:
        excel_data_V2[8][4] = CCSS_part_patronale # E37

    excel_data_V2[8][5] = excel_data_V2[8][1] * excel_data_V2[8][4] # F37

    #ligne 38
    if bulletin.Gens_de_Maison == "Oui":
        excel_data_V2[9][1] = -excel_data_V2[8][1] * 0.67 # B38
    else:
        excel_data_V2[9][1] = 0 # B38

    excel_data_V2[9][4] = CCSS_taux_reduit_part_patronale # E38
    excel_data_V2[9][5] = excel_data_V2[9][1] * excel_data_V2[9][4] # F38

    #ligne 39
    H39 = (plafond_fracais_tranche_A / 169 * excel_data_V2[2][1]) * 4 #********************************************************

    if bulletin.Exclu_ass_cho == "Oui":
        excel_data_V2[10][1] = 0 # B39
    else:
        excel_data_V2[10][1] = min(excel_data_V2[5][3], H39) # B39

    excel_data_V2[10][2] = pole_emploi_part_salariale # C39
    excel_data_V2[10][3] = excel_data_V2[10][1] * excel_data_V2[10][2] # D39
    excel_data_V2[10][4] = pole_emploi_part_patronale # E39
    excel_data_V2[10][5] = excel_data_V2[10][1] * excel_data_V2[10][4] # F39

    #ligne 40
    rounded_B31 = round(excel_data_V2[2][1])  # Arrondi de B31 à 0 décimale
    calculation = (CMRC_tranche1 / plafond_plein_CMRC) * rounded_B31 
    H40 = min(CMRC_tranche1, calculation)

    if excel_data_V2[5][3] <= H40:
        excel_data_V2[11][1] = excel_data_V2[5][3] # B40
    else:
        excel_data_V2[11][1] = H40 # B40

    excel_data_V2[11][2] = CMRC_tranche1_part_salariale # C40
    excel_data_V2[11][3] = excel_data_V2[11][1] * excel_data_V2[11][2] # D40
    excel_data_V2[11][4] = CMRC_tranche1_part_patronale # E40
    excel_data_V2[11][5] = excel_data_V2[11][1] * excel_data_V2[11][4] # F40

    #ligne 41
    H41 = H40*7

    if excel_data_V2[5][3] <= (H40 + H41):
        excel_data_V2[12][1] = excel_data_V2[5][3] - excel_data_V2[11][1] # B41
    else:
        excel_data_V2[12][1] = (H40 + H41) - excel_data_V2[11][1] # B41
    
    excel_data_V2[12][2] = CMRC_tranche2_part_salariale  # C41
    excel_data_V2[12][3] = excel_data_V2[12][1] * excel_data_V2[12][2] # D41
    excel_data_V2[12][4] = CMRC_tranche2_part_patronale # E41
    excel_data_V2[12][5] = excel_data_V2[12][1] * excel_data_V2[12][4] # F41

    #ligne 42
    B8 = "Oui" if bulletin.prevoyance else "Non"
    H42 = prevoyance_cadre_plafond_monégasque / Horaire_temps_plein_monegasque * excel_data_V2[2][1]

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[13][1] = 0 # B42
        else:
            excel_data_V2[13][1] = min(H42, excel_data_V2[5][3]) # B42
    else:
        excel_data_V2[13][1] = 0 # B42
    
    excel_data_V2[13][2] = bulletin.taux_prev_T1_sal / 100 # C42
    excel_data_V2[13][3] = excel_data_V2[13][1] * excel_data_V2[13][2]  # D42
    excel_data_V2[13][4] = bulletin.taux_prev_T1_patr / 100 # E42
    excel_data_V2[13][5] = excel_data_V2[13][1] * excel_data_V2[13][4] # F42

    #ligne 43
    H43 = H39 - H42

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[14][1] = 0 # B43
        else:
            if excel_data_V2[5][3] > H42:
                excel_data_V2[14][1] = min(excel_data_V2[5][3]-excel_data_V2[13][1],H42+H43-excel_data_V2[13][1]) # B43
            else:
                excel_data_V2[14][1] = 0 # B43
    else:
        excel_data_V2[14][1] = 0 # B43
    
    excel_data_V2[14][2] =  bulletin.taux_prev_T2_sal / 100 # C43
    excel_data_V2[14][3] =  excel_data_V2[14][1] * excel_data_V2[14][2] # D43
    excel_data_V2[14][4] =  bulletin.taux_prev_T2_patr / 100 # E43
    excel_data_V2[14][5] =  excel_data_V2[14][1] * excel_data_V2[14][4] # F43

    #ligne 44
    H46 = plafond_fracais_tranche_A / Horaire_temps_plein_monegasque * excel_data_V2[2][1]

    if B8 == "Oui":
        if bulletin.type_plafond == "FR URSSAF":
            excel_data_V2[15][1] = min(excel_data_V2[5][3],H46) # B44
        else:
            excel_data_V2[15][1] = 0 # B44
    else:
        excel_data_V2[15][1] = 0 # B44

    excel_data_V2[15][2] = bulletin.taux_prev_T1_sal / 100 # C44
    excel_data_V2[15][3] = excel_data_V2[15][1] * excel_data_V2[15][2]  # D44
    excel_data_V2[15][4] = bulletin.taux_prev_T1_patr / 100 # E44
    excel_data_V2[15][5] = excel_data_V2[15][1] * excel_data_V2[15][4] # F44

    #ligne 45
    H49 = H46*4

    if B8 == "Non":
        excel_data_V2[16][1] = 0 # B45
    else:
        if bulletin.type_plafond != "FR URSSAF":
            excel_data_V2[16][1]  = 0 # B45
        else:
            if excel_data_V2[5][3] > H46: 
                excel_data_V2[16][1] = min(excel_data_V2[5][3] - excel_data_V2[15][1], H49 - excel_data_V2[15][1]) # B45
            else:
                excel_data_V2[16][1] = 0 # B45

    excel_data_V2[16][2] = bulletin.taux_prev_T2_sal / 100 # C45
    excel_data_V2[16][3] = excel_data_V2[16][1] * excel_data_V2[16][2] # D45
    excel_data_V2[16][4] = bulletin.taux_prev_T2_patr / 100 # E45
    excel_data_V2[16][3] = excel_data_V2[16][1] * excel_data_V2[16][4] # F45


    #ligne 46
    excel_data_V2[17][1] =  excel_data_V2[5][3] # B46

    if bulletin.Taux_ATHT > 0:
        excel_data_V2[17][4] = bulletin.Taux_ATHT / 100 # E46
    else:
        excel_data_V2[17][4] = 1 / 100 # E46

    excel_data_V2[17][5] = excel_data_V2[17][1] * excel_data_V2[17][4] # F46

    #ligne 47
    excel_data_V2[18][1] = excel_data_V2[17][5] # B47

    excel_data_V2[18][4] = taxe_cotisation_AT_patronale # E47

    excel_data_V2[18][5] = excel_data_V2[18][1] * excel_data_V2[18][4] # F47

    #ligne 48
    if bulletin.CCPB_ouvrier == "Oui":
        excel_data_V2[19][1] = excel_data_V2[5][3] # B48
    else:
        excel_data_V2[19][1] = 0 # B48

    excel_data_V2[19][2] = CCPB_OUVRIER_salariale # C48

    excel_data_V2[19][3] = excel_data_V2[19][1] * excel_data_V2[19][2] # D48
    excel_data_V2[19][4] = CCPB_OUVRIER_patronale # E48
    excel_data_V2[19][5] = excel_data_V2[19][1] * excel_data_V2[19][4] # F48
    

    #ligne 49

    if bulletin.CCPB_Etam_cadre == "Oui":
        excel_data_V2[20][1] = excel_data_V2[5][3] # B49
    else:
        excel_data_V2[20][1] = 0 # B49

    excel_data_V2[20][2] = CCPB_ETAM_CADRE_salariale # C49 

    excel_data_V2[20][3] = excel_data_V2[20][1] * excel_data_V2[20][2] # D49
    excel_data_V2[20][4] = CCPB_ETAM_CADRE_patronale # E49
    excel_data_V2[20][5] = excel_data_V2[20][1] * excel_data_V2[20][4] # F49


    #ligne 51
    excel_data_V2[22][3] = 0
    for i in range(7,21):
        excel_data_V2[22][3] = excel_data_V2[22][3] + excel_data_V2[i][3] # D51

    excel_data_V2[22][5] = 0
    for i in range(7,21):
        excel_data_V2[22][5] = excel_data_V2[22][5] + excel_data_V2[i][5] # F51

    #ligne 53
    excel_data_V2[24][2] = excel_data_V2[22][3] / excel_data_V2[5][3] # C53

    excel_data_V2[24][4] = excel_data_V2[22][5] / excel_data_V2[5][3] # E53

    #ligne 56
    H50 = SMIC * 1.05

    if excel_data_V2[2][2] <= H50:
        excel_data_V2[27][1] = excel_data_V2[5][3] # B56
    else:
        excel_data_V2[27][1] = 0 # B56

    excel_data_V2[27][2] = 0.05 # C56

    excel_data_V2[27][3] = excel_data_V2[27][1] * excel_data_V2[27][2] # D56

    excel_data_V2[27][5] = excel_data_V2[27][3] # F56

    #ligne 57
    excel_data_V2[28][3] = bulletin.remboursement_transport # D57

    excel_data_V2[28][5] = excel_data_V2[28][3] # F57

    #ligne 58
    excel_data_V2[29][1] = bulletin.ticket_resto_prix # B58
    excel_data_V2[29][2] = -bulletin.ticket_resto_sal # C58
    excel_data_V2[29][3] = excel_data_V2[29][1] * excel_data_V2[29][2] # D58
    excel_data_V2[29][4] = bulletin.ticket_resto_patr # E58
    excel_data_V2[29][5] = excel_data_V2[29][4] * excel_data_V2[29][1] # F58

    #ligne 59
    B12 = "Oui" if bulletin.mutuelle else "Non"

    if B12 == "Oui":
        excel_data_V2[30][3] = -bulletin.montant_mutu_sal # D59
    else:
        excel_data_V2[30][3] = 0 # D59

    if B12 == "Oui":
        excel_data_V2[30][5] = bulletin.montant_mutu_patr # F59
    else:
        excel_data_V2[30][5] = 0 # F59

    #ligne 60
    excel_data_V2[31][3] = 20 # D60

    #ligne 61
    excel_data_V2[32][1] = bulletin.panier_prix # B61
    excel_data_V2[32][2] = bulletin.paniers_sal # C61
    excel_data_V2[32][3] = excel_data_V2[32][1] * excel_data_V2[32][2] # D61
    excel_data_V2[32][5] = excel_data_V2[32][3] # F61

    #ligne 62
    excel_data_V2[33][5] = 0
    for i in range(22,33):
        excel_data_V2[33][5] = excel_data_V2[33][5] + excel_data_V2[i][5] # F62

    #ligne 63
    somme = 0
    for i in range(24,33):
        somme = somme + excel_data_V2[i][3]

    excel_data_V2[34][3] = excel_data_V2[2][3] - excel_data_V2[22][3] + somme # D63

    #ligne 65
    excel_data_V2[36][1] = excel_data_V2[33][5] + excel_data_V2[2][3] # B65

    #ligne 66
    excel_data_V2[37][1] = excel_data_V2[36][1]*12 # B66


    print(len(excel_data_V2))
    print(len(excel_data_V2[0]))
    print(type(excel_data_V2))

    # Formatage des datas (arrondis, %, ...)
    for i in range(7, 28):
        if excel_data_V2[i][2] is not None:
            excel_data_V2[i][2] = excel_data_V2[i][2] * 100

    for i in range(7, 25):
        if excel_data_V2[i][4] is not None:
            excel_data_V2[i][4] = excel_data_V2[i][4] * 100

    for i in range(2,38):
        for j in range(1,6):
            if (excel_data_V2[i][j] is not None) and not isinstance(excel_data_V2[i][j], str):
                if (i== 2) and (j==2):
                    excel_data_V2[i][j] = f"{excel_data_V2[i][j]:.4f}" # attention ça fait que maintenant les chiffres deviennent des strings
                else:
                    excel_data_V2[i][j] = f"{excel_data_V2[i][j]:.2f}" # attention ça fait que maintenant les chiffres deviennent des strings

                if (excel_data_V2[i][j] == f"{0.0:.2f}"):
                    excel_data_V2[i][j] = " "

                excel_data_V2[i][j] = str(excel_data_V2[i][j])

                for p in range(0, len(excel_data_V2[i][j])):
                    if (excel_data_V2[i][j][p] == ".") and (len(excel_data_V2[i][j][:p]) >= 4) and (len(excel_data_V2[i][j][:p]) < 7):
                        u = p - 3
                        excel_data_V2[i][j] = excel_data_V2[i][j][:u] + " " + excel_data_V2[i][j][u:]
                                

    excel_data_V2[34][3] = excel_data_V2[34][3] + " €"
    excel_data_V2[36][1] = excel_data_V2[36][1] + " €"
    excel_data_V2[37][1] = excel_data_V2[37][1] + " €"

    for i in range(7, 28):
        if excel_data_V2[i][2] is not None and excel_data_V2[i][2] != " ":
            excel_data_V2[i][2] = str(excel_data_V2[i][2]) + "%"

    for i in range(7, 25):
        if excel_data_V2[i][4] is not None and excel_data_V2[i][4] != " ":
            excel_data_V2[i][4] = str(excel_data_V2[i][4]) + "%"

    

    prev_bulletin = Cout_to_Brut.objects.filter(id__lt=bulletin_id).order_by('-id').first()
    next_bulletin = Cout_to_Brut.objects.filter(id__gt=bulletin_id).order_by('id').first()

    if prev_bulletin is not None:
        prev_bulletin = prev_bulletin.bulletin

    if next_bulletin is not None:
        next_bulletin = next_bulletin.bulletin

    
    brut_heure = cout_vers_brut.salaire_brut / bulletin.heures_mois
    brut_heure = round(brut_heure,4)


    # Pass the formatted data to the template
    context = {'excel_data': excel_data_V2,
        "bulletin_ID" : bulletin_id,
        'bulletin': bulletin,
        'prev_bulletin': prev_bulletin,
        'next_bulletin': next_bulletin,
        "cout_vers_brut": cout_vers_brut,
        "brut_heure": brut_heure
    }

    return render(request, 'simulateur_paie/rendu_Cout_to_brut.html', context)



def edit_bulletin(request, bulletin_id):
    bulletin = Bulletin.objects.get(pk=bulletin_id)

    if request.method == "POST":  # The form uses POST for validation
        form = BulletinForm(request.POST, instance=bulletin)  # Passing an instance to avoid empty form data

        if form.is_valid():
            form.save()
            return redirect('simulateur_paie:bulletin_detail', bulletin_id)
        else:
            # If form is not valid, set section2_checked based on the form data
            section2_checked = (
                form.cleaned_data.get('prevoyance', False) or
                form.cleaned_data.get('mutuelle', False) or
                form.cleaned_data.get('taux_ccss_red') == "Oui" or
                form.cleaned_data.get('admin_SAM') == "Oui" or
                form.cleaned_data.get('Exclu_ass_cho') == "Oui" or
                form.cleaned_data.get('CCPB_ouvrier') == "Oui" or
                form.cleaned_data.get('CCPB_Etam_cadre') == "Oui" or
                form.cleaned_data.get('Gens_de_Maison') == "Oui"
            )
    else:
        form = BulletinForm(instance=bulletin)
        section2_checked = (
            bulletin.prevoyance or
            bulletin.mutuelle or
            bulletin.taux_ccss_red == "Oui" or
            bulletin.admin_SAM == "Oui" or
            bulletin.Exclu_ass_cho == "Oui" or
            bulletin.CCPB_ouvrier == "Oui" or
            bulletin.CCPB_Etam_cadre == "Oui" or
            bulletin.Gens_de_Maison == "Oui"
        )

    return render(request, 'simulateur_paie/bulletin_edit.html', {"form": form, "section2_checked": section2_checked})
    

def edit_bulletin_net_to_brut(request, bulletin_id):
    net_tobrut = Net_to_Brut.objects.get(pk=bulletin_id)
    bulletin = net_tobrut.bulletin

    if request.method == "POST":  # The form uses POST for validation
        form = BulletinForm_Net_Vers_Brut(request.POST, instance=bulletin)  # Passing an instance to avoid empty form data

        if form.is_valid():

            bulletin_res = form.save()
            B_ID = bulletin_res.id
            salaire_net_souhaité = bulletin_res.salaire_net_mois
            indemnite = bulletin_res.indemnite
            heures = bulletin_res.heures_mois
            H50 = 11.65 * 1.05

            # Hypothèse initiale : net * 1.13 pour estimer un brut de départ
            brut_estime = salaire_net_souhaité * 1.13
            taux_horaire = brut_estime/heures
            
            if indemnite == "Oui":
                i = 0.01
                while taux_horaire > H50:
                    brut_estime =  salaire_net_souhaité * (1.13-i)
                    taux_horaire = brut_estime/heures
                    i = i + 0.01

            
            max_iter=10000000  # 10 000 000

            for i in range(max_iter):
            
                salaire_net_simulé = simu_net_avec_brut_estime(brut_estime,B_ID)
                
                # Comparer le net simulé au net souhaité
                différence = salaire_net_souhaité - salaire_net_simulé
                
                # Si la différence est inférieure à la tolérance, on a trouvé une solution
                if abs(différence) == 0:
                    break
                
                # Ajuster le brut estimé en fonction de la différence
                brut_estime += différence * 0.1  # Ajustement itératif léger

            brut_estime = round(brut_estime,1)


            net_tobrut.salaire_net_mois = salaire_net_souhaité
            net_tobrut.salaire_brut = brut_estime 
            net_tobrut.bulletin = form.save()
            net_tobrut.save()
            
            return redirect('simulateur_paie:bulletin_net_to_brut', bulletin_id)
        
        else:
            # If form is not valid, set section2_checked based on the form data
            section2_checked = (
                form.cleaned_data.get('prevoyance', False) or
                form.cleaned_data.get('mutuelle', False) or
                form.cleaned_data.get('taux_ccss_red') == "Oui" or
                form.cleaned_data.get('admin_SAM') == "Oui" or
                form.cleaned_data.get('Exclu_ass_cho') == "Oui" or
                form.cleaned_data.get('CCPB_ouvrier') == "Oui" or
                form.cleaned_data.get('CCPB_Etam_cadre') == "Oui" or
                form.cleaned_data.get('Gens_de_Maison') == "Oui"
            )
    else:
        form = BulletinForm_Net_Vers_Brut(instance=bulletin)
        section2_checked = (
            bulletin.prevoyance or
            bulletin.mutuelle or
            bulletin.taux_ccss_red == "Oui" or
            bulletin.admin_SAM == "Oui" or
            bulletin.Exclu_ass_cho == "Oui" or
            bulletin.CCPB_ouvrier == "Oui" or
            bulletin.CCPB_Etam_cadre == "Oui" or
            bulletin.Gens_de_Maison == "Oui"
        )

    return render(request, 'simulateur_paie/simu_paie_form_Net_to_Brut_edit.html', {"form": form, "section2_checked": section2_checked})


def edit_bulletin_cout_to_brut(request, bulletin_id):
    cout_tobrut = Cout_to_Brut.objects.get(pk=bulletin_id)
    bulletin = cout_tobrut.bulletin

    if request.method == "POST":  # The form uses POST for validation
        form = BulletinForm_Cout_Vers_Brut(request.POST, instance=bulletin)  # Passing an instance to avoid empty form data

        if form.is_valid():

            bulletin_res = form.save()
            B_ID = bulletin_res.id

            cout_souhaité = bulletin_res.cout_mois # ancien nom : salaire_net_souhaité
            indemnite = bulletin_res.indemnite
            heures = bulletin_res.heures_mois
            H50 = 11.65 * 1.05

            # Hypothèse initiale : cout * 0.714 pour estimer un brut de départ
            brut_estime = cout_souhaité * 0.714
            taux_horaire = brut_estime/heures
            
            if indemnite == "Oui":
                i = 0.01
                while taux_horaire > H50:
                    brut_estime =  cout_souhaité * (0.714-i)
                    taux_horaire = brut_estime/heures
                    i = i + 0.01

            
            max_iter=10000000  # 10 000 000

            for i in range(max_iter):
            
                cout_simulé = simu_net_avec_cout_estime(brut_estime,B_ID)
                
                # Comparer le net simulé au net souhaité
                différence = cout_souhaité - cout_simulé
                
                # Si la différence est inférieure à la tolérance, on a trouvé une solution
                if abs(différence) == 0:
                    break
                
                # Ajuster le brut estimé en fonction de la différence
                brut_estime += différence * 0.1  # Ajustement itératif léger

            brut_estime = round(brut_estime,1)

            cout_tobrut.cout_mois = cout_souhaité
            cout_tobrut.salaire_brut = brut_estime 
            cout_tobrut.bulletin = form.save()
            cout_tobrut.save()
            
            return redirect('simulateur_paie:bulletin_cout_to_brut', bulletin_id)
        
        else:
            # If form is not valid, set section2_checked based on the form data
            section2_checked = (
                form.cleaned_data.get('prevoyance', False) or
                form.cleaned_data.get('mutuelle', False) or
                form.cleaned_data.get('taux_ccss_red') == "Oui" or
                form.cleaned_data.get('admin_SAM') == "Oui" or
                form.cleaned_data.get('Exclu_ass_cho') == "Oui" or
                form.cleaned_data.get('CCPB_ouvrier') == "Oui" or
                form.cleaned_data.get('CCPB_Etam_cadre') == "Oui" or
                form.cleaned_data.get('Gens_de_Maison') == "Oui"
            )
    else:
        form = BulletinForm_Cout_Vers_Brut(instance=bulletin)
        section2_checked = (
            bulletin.prevoyance or
            bulletin.mutuelle or
            bulletin.taux_ccss_red == "Oui" or
            bulletin.admin_SAM == "Oui" or
            bulletin.Exclu_ass_cho == "Oui" or
            bulletin.CCPB_ouvrier == "Oui" or
            bulletin.CCPB_Etam_cadre == "Oui" or
            bulletin.Gens_de_Maison == "Oui"
        )

    return render(request, 'simulateur_paie/simu_paie_form_Cout_to_Brut_edit.html', {"form": form, "section2_checked": section2_checked})



def remove_bulletin(request, bulletin_id):
    try:
        book = Bulletin.objects.get(pk = bulletin_id)
        book.delete()
    except Bulletin.DoesNotExist:
        return redirect("simulateur_paie:historique")
    return redirect("simulateur_paie:historique")


def remove_bulletin_net_to_brut(request, bulletin_id):
    try:
        book = Net_to_Brut.objects.get(pk = bulletin_id)
        book.delete()
    except Net_to_Brut.DoesNotExist:
        return redirect("simulateur_paie:historique_net_to_brut")
    return redirect("simulateur_paie:historique_net_to_brut")


def remove_bulletin_cout_to_brut(request, bulletin_id):
    try:
        book = Cout_to_Brut.objects.get(pk = bulletin_id)
        book.delete()
    except Cout_to_Brut.DoesNotExist:
        return redirect("simulateur_paie:historique_cout_to_brut")
    return redirect("simulateur_paie:historique_cout_to_brut")
