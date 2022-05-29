# Author: Boris Mauricette <teddy_boomer@yahoo.fr>
# License: GNU GPL v3
# v0.2

import pandas as pd
import tkinter as tk
from tkinter.filedialog import askopenfilename
import time
import os
import os.path
from pandas_ods_reader import read_ods
from functools import reduce
import matplotlib.pyplot as plt
from mpl_toolkits.axes_grid1.axes_divider import make_axes_area_auto_adjustable
import json
import numpy as np

# script d'extraction des moyennes élève sio2 pour:
#  - construire le graphe
#  - reformer les résultats vers le modèle académique
# distribué en l'état; assurez-vous qu'il fait ce que vous attendez!

elevesFile = askopenfilename(initialdir="", # os.environ['HOME']
                             title="Fichier listing élèves 2e année",
                             filetypes=[("ODS", "*.ods")])
eleves = pd.read_excel(elevesFile, engine="odf",
                       parse_dates=[1], sheet_name=None)

# TODO: proposer un choix de filière
with open("filieres.json", "r") as F: 
    matieresData = json.load(F)["SIO"]
disciplines = matieresData["disciplines"]
disc_titles = matieresData["disc_titles"]
TI = dict(zip(disciplines, disc_titles))
toplot = disciplines[0:8] # ou [:-1]

# fonctions techniques
def is_discipline(e):
    """tester si e est le titre d'une des disciplines à conserver
    """
    return reduce(lambda a,b: a or b, [f in e for f in disciplines])

def is_toplot(e):
    """test si e est le titre d'une des disciplines à tracer
    """
    return reduce(lambda a,b: a or b, [f in e for f in toplot])

def get_rang(s):
    """renvoie le couple rg/nb eleves depuis la chaîne 'rg /nb'
    rtype: list of int
    """
    return tuple(map(int, s.split(' /')))

def get_title(e):
    """capturer le nom précis puis le vrai titre associé
    e est un des intitulés de matière pronote
    """
    g = [f for f in disciplines if f in e][0]
    return TI[g]

# astuce ordinale pour toujours ordonner les matières selon leur position dans
# disciplines
# à envisager sur MultiIndex pour NDRC disc×semestre
ORDRE = dict((e,i+1) for i,e in enumerate(disciplines))
def get_ordre(e):
    """capturer l'index précis puis le vrai titre associé
    e est un des intitulés de matière pronote
    """
    g = [f for f in disciplines if f in e][0]
    return ORDRE[g]

notations = ["Sem1", "Sem2", "Année"]
# colonnes
# ['Disciplines', 'Notation', 'Rang', 'Moy Eleve', 'Moy Classe', '<8',
#       'Entre 8 et 12', '>=12', 'Appréciations des professeurs']

def makeLivret(F, eleve, fichier_tableur, logfile, bts='SIO'):
    """créer le graphe et la feuille de tableur pour eleve

    Paramètres:
    -----------
    F: writer xlsx
    eleve: nom complet de l'élève 'NOM Prénom' extrait de pronote
    logfile: buffer d'enregistrement de la trace

    actions:
    une feuille de tableur nommée eleve ajoutée à F,
    un graphe eleve.png ajouté dans le dossier
    """
    try:
        df = pd.read_excel(fichier_tableur, sheet_name=eleve)
        # 2 rows à supprimer en 2e année:
        DROP = matieresData["to_drop"] # ["AT. PROFESSION.", "CYBER.SERV.INF."]
        
        # dfe2: dataframe eleve 2e année
        # données pivotées
        dfe2 = df.pivot(index='Disciplines', columns='Notation',
                        values='Moy Eleve').drop(DROP, axis=0).replace(to_replace=['Abs'], value=np.NaN)
        dfc = df.pivot(index='Disciplines', columns='Notation',
                       values='Moy Classe').drop(DROP, axis=0)      
        dfa = df.pivot(index='Disciplines', columns='Notation',
                       values='Appréciations des professeurs').drop(DROP, axis=0)

        # matieres du graphique
        matieres = [ e for e in dfa.index if is_toplot(e)]
        mat2title = dict([(e, get_title(e)) for e in matieres])

        # liste complète des matières
        matFull =  [ e for e in dfa.index if is_discipline(e)]
        full2title = dict([(e, get_title(e)) for e in matFull])

        # var des tuple matiere, indice pour tri
        prn_o = sorted([(e, get_ordre(e)) for e in matieres], key=lambda x: x[1])
        prn = list(zip(*prn_o))[0]
        # idem full matieres
        prnf_o = sorted([(e, get_ordre(e)) for e in matFull], key=lambda x: x[1])
        prnf = list(zip(*prnf_o))[0]
        
        dfe2.insert(3, 'Appréciations des professeur(e)s', dfa.get('Année'))
        # réordonner les colonnes
        dfe2 = dfe2.reindex(notations + ['Appréciations des professeur(e)s'],
                            axis=1)

        # export de la feuille de tableur    
        livret = dfe2.loc[matFull].reindex(prnf).rename(index=full2title)
        livret.to_excel(F, sheet_name=eleve, startrow=3)

        # construction du graphique - moyennes élève me, moyennes classe mc
        makeGraphique(dfe2, dfc, eleve, matieres, mat2title, prn, outdir, "SIO")
        print("OK", file=logfile)
    except ValueError:
        print(f"{eleve} en erreur", file=logfile)
    except KeyError:
        print(f"{eleve} en erreur - clé manquante", file=logfile)
    except TypeError:
        print(f"{eleve} en erreur - type error", file=logfile)


# Traitement graphique d'une dataFrame
def makeGraphique(dfe2, dfc, eleve, matieres, matieresTitres,
                  mat_ordered, dir="", bts='SIO'):
    """créer le graphique et l'exporter

    Paramètres:
    -----------

    dfe2: dataFrame élève
    dfc: dataFrame classe
    matières: liste des matières pronotes
    matièresTitres: dict d'association vers les titres lisibles
    mat_ordered: matières dans l'ordre attendu par le modèle
    dir: dossier d'enregistrement des images
    bts: filière (seulement 'SIO' pour l'instant)
    """
    me = dfe2.get('Année').filter(items=matieres) # .fillna(value=0)
    mc = dfc.get('Année').fillna(value=10).filter(items=matieres)
    profil = (me.div(mc)*10).round().reindex(mat_ordered).rename(index=matieresTitres)
    plt.figure(figsize=(8,9))
    ax = plt.subplot(111)
    ax.axhline(y=10, color="k", linestyle='-') 
    glabs = [get_title(e) for e in me.index]
    P = profil.plot(grid=True, ylim=(0,20),
                    xticks=range(len(me.index)),
                    yticks=range(21), style='rx-')
    plt.xticks(rotation=90)
    plt.title(f"Profil de: {eleve}")
    make_axes_area_auto_adjustable(ax)
    # plt.show()
    plt.savefig(os.path.join(dir, f"{eleve}.png"))
    plt.close()

    
for s in eleves: # onglets
    print(s)
    outdir = os.path.join(os.path.dirname(elevesFile), f"export_{s}")
    fichier_tableur = os.path.join(outdir, "eleves2_pronote.xlsx")
    sortie = os.path.join(outdir, "livrets.xlsx")
    log = os.path.join(outdir,
                   f"log_livrets-{time.strftime('%d%m%Y-%H-%M-%S')}.log")
    LOG = open(log, "w")
    if os.path.exists(fichier_tableur):
        with pd.ExcelWriter(sortie, engine='xlsxwriter',
                            datetime_format='dd/mm/yyyy') as writer:
            # engine_kwargs={'options': {'datetime_format': 'dd/mm/yyyy'}}
            for e in eleves[s].iterrows(): # e[0]: index, e[1] row
                nom = e[1]['Élève']
                print(f"traitement {nom}:", end="", file=LOG)
                # nom et date naissance sur les 2 1eres lignes         
                df = pd.DataFrame(e[1]).T
                df.to_excel(writer, sheet_name=nom)
                # reste du traitement
                makeLivret(writer, nom, fichier_tableur, LOG)
        
    else:
        print(f"pas de données pronote pour l'onglet **{s}**")
    LOG.close()
