import xlwings as xw
from matplotlib import pyplot as plt
from scipy.optimize import curve_fit
from scipy.stats import linregress
import numpy as np



###########################
#définir les données utilisées
Nom_du_fichier_XL = f'scope_geo_{0}.csv'
Nom_de_la_feuille = 'scope_geo_0'
X_data = 'A3:A1002'
Y_data = 'B3:B1002'
###########################
#définir le graphique
Titre = "Approximation gaussienne d'une fibre hautement multimodes"
Titre_X = 'Distance (pixels)'
Titre_Y = 'Gray value'
Nom_des_données = 'Données expérimentales'
Nom_de_la_courbe_de_régression = 'Fonction gaussienne'
Taille_des_points = 1
###########################
#définir les bounds des variables
A = [400, 800]
B = [-np.inf, np.inf]
C = [-np.inf, np.inf]
D = [-np.inf, np.inf]
E = [-np.inf, np.inf]
###########################
#Définir la fonction utilisé
#Gaussian: C*(1/(B*np.sqrt(2*np.pi)))*np.exp(-((x - A)**2) / (2 * B**2)) + D
#Sinus: A*np.sin(B*(x-C))+D
#Droite: A*x+B
#
#
#vous n'êtes pas obligé d'utilisé tous les paramètres
def fonction(x, A, B, C, D, E): 
    return C*(1/(B*np.sqrt(2*np.pi)))*np.exp(-((x - A)**2) / (2 * B**2)) + D

###########################



#Prendre les données d'un document excell
wb = xw.Book(Nom_du_fichier_XL)
sht = wb.sheets[Nom_de_la_feuille]
X = sht.range('A:A').value
Y = sht.range('B:B').value
print(X)

#Déterminer les paramètres du curve_fit
paramètres_fit, autre_information = curve_fit(fonction, X, Y, 
    bounds= ([A[0] , B[0], C[0], D[0], E[0]],[A[1], B[1], C[1], D[1], E[1]]))
A_fit, B_fit, C_fit, D_fit, E_fit = paramètres_fit

#Tracer la courbe de régression
Y_fit = fonction(X, A_fit, B_fit, C_fit, D_fit, E_fit)

#Print les données de la régression
#r2_score(Y_fit, Y)
#print(f'r2_score = {r2_score}')
r = linregress(Y, Y_fit)
print(f'r = {round(r[2], 3)} and r**2 = {round((r[2]**2),3)}')
print(f'A = {A_fit}, B = {B_fit}, C = {C_fit}, D = {D_fit}, E = {E_fit}')

#Tracer le graphique
fig, ax = plt.subplots()
ax.tick_params(axis='x', direction='in')
ax.tick_params(axis='y', direction='in')
plt.plot(X, Y_fit, '-r', label=Nom_de_la_courbe_de_régression)
plt.scatter(X, Y, label=Nom_des_données, s= Taille_des_points)
plt.legend() 
plt.title(Titre)
plt.xlabel(Titre_X)
plt.ylabel(Titre_Y)
plt.show()
wb.close()
