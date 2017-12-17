# -*- coding: utf-8 -*-

import os
import xlsconverter
from projectreader import clear, session


# hello world,
clear()
print("\nCe script vous permet de charger des projets etudiants, de les noter et les commenter, puis de les retranscrire dans le fichier excel de section et de les mettre en ligne sous forme d'une page web.\n")
session_0 = session("NOM","./session/session.json")
session_0.selection()
session_0.run()
