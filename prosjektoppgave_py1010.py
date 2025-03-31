#Prosjektoppgaven

import numpy as np
import matplotlib.pyplot as plt
import pandas as pd

#del a)
#Leser inn Excel-filen
support_fil = pd.read_excel("support_uke_24.xlsx") 
#Gjør innholdet i kolonnene til Arrays
u_dag = support_fil["Ukedag"].values
kl_slett = support_fil["Klokkeslett"].values
varighet = support_fil["Varighet"].values
score = support_fil["Tilfredshet"].values

#del b)

#definerer en prosedyre for å telle antall samtaler pr ukedag, og oppretter variabler pr dag, samt en tom liste-Dagene
def antall_pr_ukedag (u_dag):
    Mandag = 0
    Tirsdag = 0
    Onsdag = 0
    Torsdag = 0
    Fredag = 0
    Dagene = []
#for-løkke, som går gjennom hvert element i u_dag, og sjekker om den finnes i listen Dagene, og legger den til i listen hvis ikke
    for unike_d in u_dag:
        if unike_d  not in Dagene:
            Dagene.append(unike_d)
#ny for-løkke som også går gjennom hvert element i listen, og sjekker Navnet, hvis det stemmer overens med strengen jeg har laget, så legges det til 1 i variabelen jeg definerte tidligere
    for dag in u_dag:
        if dag == "Mandag":
            Mandag +=1
        elif dag == "Tirsdag":
            Tirsdag +=1
        elif dag == "Onsdag":
            Onsdag +=1
        elif dag == "Torsdag":
            Torsdag +=1
        elif dag == "Fredag":
            Fredag +=1
#oppretter ny variabel som innholder alle de andre variablene med antall samtaler pr ukedag    
    Antall_pr_dag = Mandag,Tirsdag,Onsdag,Torsdag,Fredag
#Lager stolpediagram, hvor x-linjen består av dagvariabelen, og y linjen viser antall samtaler pr dag        
    plt.bar(Dagene, Antall_pr_dag)
    plt.title("Antall samtaler pr dag")
    plt.ylabel("Antall samtaler")
    plt.show()
#Kaller prosedyren med u_dag som variabel
antall_pr_ukedag(u_dag)

#del c)
#Oppretter prosedyre, som sjekker max verdi på varighetkolonnen, deler opp innholdet i en liste med 3 elementer vha split-funksjonen på begge kolonne i tiden, legger hver av disse i en variabel
def skriv_ut_lengst_samtale(varighet):
    lengst_tid = np.max(varighet)
    oppdelt_max = lengst_tid.split(":")
    lengst_time= oppdelt_max[0]
    lengst_min = oppdelt_max[1]
    lengst_sek = oppdelt_max[2]
#if-setning som sjekker om variabelen = "00", hvis den ikke er det printes det ut til terminalen, dette for å unngå at teksten sier feks 00 timer og x minutter    
    if lengst_time != "00":
        print ("Den lengste samtalen i uke 24 varte i", lengst_time, "timer,", lengst_min,"minutter og", lengst_sek, "sekunder")
    elif lengst_min!= "00":
        print("Den lengste samtalen i uke 24 varte i", lengst_min,"minutter og", lengst_sek, "sekunder")
    else:
        print("Den lengste samtalen i uke 24 varte i", lengst_sek, "sekunder")
#Kaller prosedyren med argumentet varighet-kolonnen        
skriv_ut_lengst_samtale(varighet)

#ny prosdyre, helt lik som over bortsett fra at den sjekker minste verdi på varighetskolonnen
def skriv_ut_kortest_samtale(varighet):
    minst_tid = np.min(varighet)
    oppdelt_min = minst_tid.split(":")
    kortest_time = oppdelt_min[0]
    kortest_min = oppdelt_min[1]
    kortest_sek = oppdelt_min[2]
    
    if kortest_time != "00":
        print("Den korteste samtalen i uke 24 varte i ", kortest_time, "timer,", kortest_min, "minutter og ", kortest_sek, "sekunder")
    elif kortest_min!= "00":
        print("Den korteste samtalen i uke 24 varte i", kortest_min,"minutter og", kortest_sek, "sekunder")
    else:
        print("Den lengste samtalen i uke 24 varte i", kortest_sek, "sekunder")
skriv_ut_kortest_samtale(varighet)
    
#del d)
#opretter en funksjon og variabler for tid i sekunder, minutter og timer som foreløpig = 0
def gjennomsnitt_samtale(varighet):
    tid_i_sek = 0
    tid_i_min = 0
    tid_i_timer = 0
#for løkke som for hvert element i varighet kolonnen, splitter opp opp til en liste av 3 elementer, og plusser på verdien i hver variabel, slik at produktet blir totalt antall timer, minutter, sekunder    
    for tid in varighet:
        oppdelt_tid = tid.split(":")
        timer = oppdelt_tid[0]
        minutter = oppdelt_tid[1]
        sekunder = oppdelt_tid[2]
        tid_i_sek += int(sekunder)
        tid_i_min += int(minutter)
        tid_i_timer += int(timer)
#Omregner minutter og timer til sekunder ved å gange med henholdsvis 60 og 3600          
    total_min_i_sek = tid_i_min *60   
    total_timer_i_sek = tid_i_timer *3600
#regner ut snitt samtalen ved å legge sammen totalt antall sekunder, og dele det på antall elementer i varighetskolonnen, som hentes ved len funksjonen.   
    snitt_samtale = (total_timer_i_sek + total_min_i_sek + tid_i_sek) / len(varighet)
    snitt_i_min = snitt_samtale//60 # deler gjennomsnittssekunder regnet over med //, som gir heltallet og kutter resten - har da snitt-minutter brukt
    rest_sek = snitt_samtale %60 # %funksjonen finner restene av snitt-sekundende etter snitt-sekundene er delt på 60 - og 
#returnerer avrundede snitt i minutter og sekunder i en liste med 2 elementer    
    return round(snitt_i_min), round(rest_sek) 
#kaller på funksjonen, med varighetkolonnen som argument, og deretter printer jeg ut begge elementene fra listen som blir reutnert av funksjonen, med passende strenger 
snitt_uke_24 = gjennomsnitt_samtale(varighet)
print("Gjennomsnittlig samtaletid i uke 24 var", snitt_uke_24[0], "minutter og", snitt_uke_24[1], "sekunder")



#del e)
#oppretter funksjon med kl_slett som parameter, og opretter 4 bolke-variabler med verdien 0
def sektor_bolker(kl_slett):
    bolk_1 = 0
    bolk_2 = 0
    bolk_3 = 0
    bolk_4 = 0
#for-løkke, som deler opp hvert element i kl_slett til en liste, og tar det første elementer og legger det til i en kl_time variabel    
    for kl in kl_slett:
        oppdelt_kl = kl.split(":")
        kl_time = oppdelt_kl[0]
#Sjekker verdien i kl_time variabelen, og det avgjør da hvilken bolke-variabel den skal legge til 1 i         
        if kl_time == "08" or kl_time =="09":
            bolk_1 += 1
        elif kl_time == "10" or kl_time == "11":
            bolk_2 +=1
        elif kl_time == "12" or kl_time == "13":
            bolk_3 +=1
        elif kl_time == "14" or kl_time == "15":
            bolk_4 +=1
#Opretter en liste, bolker, som består av bolkevariabelene. Opretter en variabel for totalt bolker ved å bruke sum-funksjonen av kl_time-listen. Opretter liste med navn på bolkene    
    bolker = [bolk_1,bolk_2,bolk_3,bolk_4]
    totalt_bolker = sum(bolker)
    bolkerNavn = ["kl.08-10", "kl.10-12", "kl.12-14", "kl.14-16"]
#Oppretter sektordiagram med bolker-listen som datakilde, og labels fra BolkerNavn listen. Returnerer bolker-lsiten    
    plt.figure()
    plt.pie(bolker, labels = bolkerNavn, autopct= "%d%%",startangle=90)
    plt.pie
    plt.title("Fordeling av samtaler over ulike tidsrom uke 24")
    plt.show()
    return(bolker)
#Kaller på funksjonen med kl_slett kolonnen som argument, og lagrer det til variabel med uke_24 navn
bolker_uke_24 = sektor_bolker(kl_slett)

#del f)
#Opretter funksjon og variabler med verdien 0
def beregn_nps_score(score):
    antall_svar = 0
    pos_svar = 0
    neg_svar = 0 
    noy_svar = 0
#For-løkke, som sjekker hvert element i score kolonnen, og plusser på antall_svar variabelen, og i variabel for positiv, nøytral eller negativt svar, dersom svaret er i rangen som oppgitt
    for svar in score:
        if svar in range(1,7):
            antall_svar+=1
            neg_svar +=1
        elif svar in range (7,9):
            antall_svar +=1
            noy_svar +=1
        elif svar in range(9,11):
            antall_svar +=1
            pos_svar +=1
#regner ut prosenten av negativ, nøytral og positivt svar - legger det i en variabel. Skriver deretter ut svarene med avrundet til 1 desimal med round(x,1)funksjon            
    pros_neg = neg_svar/antall_svar*100
    pros_pos = pos_svar/antall_svar*100
    pros_noy = noy_svar/antall_svar*100
    print("Det kom tilbake",round(pros_neg,1), "% negative svar (1-6)",round(pros_noy,1),"% nøytrale svar(7-8), og" ,round(pros_pos,1),"% positive svar")
#tar prosenten positive svar, minus prosenten negative svar, og lagrer det i NPS-variabel. Printer deretter ut avrundet NPS score til en desimal. Og returnerer avrundet NPS score    
    NPS = pros_pos - pros_neg
    print("NPS score=",round(NPS,1)) 
    return(round(NPS,1))
#kaller på funksjonen med score listen som argument    
uke_24 = beregn_nps_score(score)

