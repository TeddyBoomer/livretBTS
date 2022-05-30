# Author: Boris Mauricette <teddy_boomer@yahoo.fr>
# License: GNU GPL v3

import pandas as pd
import os.path
import os
from functools import reduce
from tkinter.filedialog import askopenfilename

# appliquer à un fichier listing d'élèves type eleves.ods
# mettre la nature du bts dans le nom de chaque onglet - ex: SIO2, SIO2alt
elevesFile = askopenfilename(initialdir="", #os.environ['HOME']
                             title="Fichier listing élèves 2e année",
                             filetypes=[("ODS", "*.ods")])
eleves = pd.read_excel(elevesFile, engine="odf", sheet_name=None,
                       usecols=[0, 1], parse_dates=[1])
for s in eleves: # parcours des onglets
    outdir = os.path.join(os.path.dirname(elevesFile), f"export_{s}") 
    try:
        os.mkdir(outdir)
    except FileExistsError:
        pass
    with pd.ExcelWriter(os.path.join(outdir, f"eleves_{s}_pronote.xlsx"),
                        engine='xlsxwriter',
                        datetime_format='dd/mm/yyyy') as writer:
        # engine_kwargs={'options': {'default_date_format': 'dd/mm/yyyy'}}
        for e in eleves[s].iterrows(): # e[0]: index, e[1] row
            nom = e[1]['Élève']
            print(f"traitement {nom}")
            tmp = pd.DataFrame(e[1])
            tmp.to_excel(writer, nom)
