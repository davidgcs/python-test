from openpyxl import Workbook, load_workbook

# Si quieres VOLCAR el Excel original y añadir lo nuevo, descomenta las 3 lineas siguientes:
# wb = load_workbook('Asignaturas_Medicina_Formateado.xlsx')
# ws = wb.active
# ws_new = wb.copy_worksheet(ws)
# O (más simple) sigue generando toda la tabla nueva:

wb = Workbook()
ws = wb.active
ws.title = "Sheet1"

# Cabecera exacta:
ws.append([
    "Asignatura", "Título", "Autores", "Edición / Año", "Editorial", "Notas"
])

# ----------- DATOS EXISTENTES (de tu archivo) -----------
datos_originales = [
    # (Solo los 3 primeros títulos de muestra, el resto igual, debes copiar todos los tuyos)
    ["Anatomía Funcional i Embriologia de l'Aparell Locomotor", "Benninghoff y Drenckhahn, compendio de anatomía", "Drenckhahn D, Waschke J, Negrete JH", "2010", "Médica Panamericana", ""],
    ["", "Feneis: nomenclatura anatómica ilustrada", "Dauber W, Feneis H, Spitzer G", "2007 (5a ed.)", "Masson", ""],
    ["", "Gray Anatomía para estudiantes", "Drake RL, Vogl W, Mitchell AWM", "2015 (3ª ed.)", "Elsevier", ""],
    # ... AÑADE AQUÍ EL RESTO DE TUS FILAS EXISTENTES, igual que aparecen en tu Excel ...
    ["Biología Celular y del Desarrollo","Molecular Biology of the Cell","Alberts B, et al.","2014 (6th ed.)","Garland Science","También en castellano"],
    # ... etc ...
    ["Bioestadística, Epidemiología e Investigación", "An Introduction to Medical Statistics", "Bland M", "2015 (4ª ed.)", "Oxford", ""],
    # ... ETCÉTERA HASTA EL FINAL DE TU ARCHIVO ...
]

for fila in datos_originales:
    ws.append(fila)

# ----------- FILAS NUEVAS DE TU TEXTO -----------
# Biofísica Médica General
biofisica = [
    ["Biofísica Médica General", "Física para las ciencias de la vida", "Cromer AH", "1981 (2ª ed.)", "Reverté", ""],
    ["", "Física en la ciencia y en la industria", "Cromer AH", "1986", "Reverté", ""],
    ["", "Física y biofísica: radiaciones", "Dutreix J, Desgrez A, Bok B", "1980", "AC", ""],
    ["", "Fisiología humana", "Fernández-Tresguerres JA... [et al.]", "2010 (4ª ed.)", "McGraw-Hill Interamericana Editores", ""],
    ["", "Biofísica", "Frumento AS", "1995 (3ª ed.)", "Mosby/Doyma Libros", ""],
    ["", "Física", "Kane JW, Sternheim MM", "1989 (2ª ed.)", "Reverté", ""],
    ["", "Física biológica: energía, información, vida", "Nelson PC", "2005", "Reverté", ""],
    ["", "Cellular biophysics", "Weiss TF", "1996", "The MIT Press", "2v."],
    ["", "Biofísica", "Yushimito Rubiños L", "2007", "Manual Moderno", ""],
    ["", "Física para la ciencia y la tecnología", "Tipler PA, Mosca G", "2010 (6a ed.)", "Reverté", ""],
    ["", "Física para ciencias de la vida", "David Jou Mirabent, Josep Enric Llebot Rabagliati, Carlos Pérez García", "2009 (2a ed.)", "McGraw-Hill/Interamericana de España", ""],
    ["", "Biological physics: energy, information, life", "Nelson PC, Marko Radosavljevic, Sarina Bromberg", "2014 (5th print.)", "W.H. Freeman", ""],
]

for fila in biofisica:
    ws.append(fila)

# Biología Celular
biologia_celular = [
    ["Biología Celular", "Biología molecular de la célula", "Alberts B, et al.", "2016 (6ª ed.)", "Omega", "También 6a ed. (2015) en inglés"],
    ["", "El Mundo de la célula", "Becker B, et al.", "2007 (6a edición)", "Addison Wesley", "También 9a ed. (2016) en inglés"],
    ["", "La célula", "Cooper GM, Hausman RM", "2017 (7a edición)", "Marbán", "También 7a ed. (2016) en inglés"],
    ["", "Biología celular y molecular", "Lodish H, et al.", "2016 (7a edición)", "Médica Panamericana", "También 8a ed. (2016) en inglés"],
    ["", "Cell biology", "Pollard TD, Earnshaw WC, Lippincott-Schwartz J, Johnson GT", "2016 (3th. ed.)", "Saunders Elsevier", ""],
    ["", "Physical biology of the cell", "Phillips R, Kondev J, Theriot J", "2013 (2nd ed.)", "Garland Science", ""],
    ["", "Molecular Cell Biology", "Lodish H, et al.", "2016 (8th edition)", "MacMillan Learning", ""],
    ["", "Molecular Biology of the Cell", "Alberts B, et al.", "2022 (7th edition)", "W.W. Norton", ""],
    ["", "Essential Cell Biology", "Alberts B, et al.", "2013 (4th edition)", "Garland Science", ""],
]

for fila in biologia_celular:
    ws.append(fila)

# Biología Molecular
biologia_molecular = [
    ["Biología Molecular", "Biología molecular de la célula", "Alberts B, Wilson J, Hunt T", "2016 (6ª ed.)", "Omega", "También en inglés"],
    ["", "Bioquímica médica", "Baynes JW, Dominiczak MH", "2015 (4ª ed.)", "Elsevier", "También en inglés"],
    ["", "El mundo de la célula", "Becker WM, Kleinsmith LH, Hardin J", "2007 (6ª ed.)", "Addison Wesley", "También 9a ed. en inglés"],
    ["", "Biologia", "Campbell NA, Reece JB", "2007 (7ª ed.)", "Médica Panamericana", "También 10a ed. en inglés"],
    ["", "Bioquímica ilustrada: bioquímica y biología molecular en la era posgenómica", "Campbell PN, Smith AD, Peters TJ", "2006 (5ª ed.)", "Masson", ""],
    ["", "Bioquímica: libro de texto con aplicaciones clínicas", "Devlin TM", "2015 (4ª ed.)", "Reverté", "También 7a ed. en inglés"],
    ["", "Bioquímica", "Ferrier DR, Harvey RA", "2014 (6ª ed.)", "Wolters Kluwer Health/Lippincott Williams & Wilkins", "También 6a ed. en inglés"],
    ["", "Lewin’s genes XI", "Krebs JE, Goldstein ES, Kilpatrick ST", "2014", "Jones & Bartlett Learning", ""],
    ["", "Lewin’s essential genes", "Krebs JE, Goldstein ES, Kilpatrick ST", "2013 (3rd ed.)", "Jones and Bartlett Publishers", "También 2a ed. en castellà"],
    ["", "Molecular cell biology", "Lodish H, et al.", "2016 (8th. ed.)", "Freeman", ""],
    ["", "Bioquímica médica básica: un enfoque clínico", "Lieberman M, Marks A, Peet A", "2013 (4a ed.)", "Wolters Kluwer/Lippincott Williams & Wilkins", "También 4a ed. en inglés"],
    ["", "Bioquímica", "Mathews CHK, et al.", "2013 (4ª ed.)", "Pearson Educación", "También 4a ed. en inglés"],
    ["", "Bioquímica: las bases moleculares de la vida", "Mckee T, Mckee JR", "2014 (5ª ed.)", "McGraw-Hill/Interamericana", ""],
    ["", "Bioquímica con aplicaciones clínicas", "Stryer L, Berg JM, Tymoczko JL", "2015 (7ª ed.)", "Reverté", "También 8a ed. en inglés"],
    ["", "Bioquímica", "Voet D, Voet JG", "2006 (3ª ed.)", "Médica panamericana", "También 4a ed. en inglés"],
    ["", "Molecular biology of the gene", "Watson JD, et al.", "2013 (7th ed.)", "Pearson", "También 7a ed. en castellà"],
    ["", "Biochemistry", "Jeremy M. Berg, John L. Tymoczko, Gregory J. Gatto, Jr., Lubert Stryer", "2015 (8th edition)", "", ""],
    ["", "NCBI Home Page", "", "", "", "Bases de dades"],
    ["", "PDB lite", "", "", "", "Bases de dades"],
    ["", "The JASPAR Database", "", "", "", "Bases de dades per analitzar regions reguladores gèniques"],
    ["", "UNIPROT", "", "", "", "Bases de dades"],
    ["", "The Medical Biochemistry Page", "", "", "", "Pàgina web"],
    ["", "http://www.ncbi.nlm.nih.gov/BLAST/", "", "", "", "Pàgina web"],
    ["", "GTEX Portal", "", "", "", "https://gtexportal.org/home/"],
    ["", "Protter", "", "", "", "http://wlab.ethz.ch/protter/start/"]
]

for fila in biologia_molecular:
    ws.append(fila)

# Bioquímica Básica (vacía)
ws.append(["Bioquímica Básica", "", "", "", "", "Nada aún"])

# Histología Humana
histologia = [
    ["Histología Humana", "Geneser histología : 4a. edición", "Brüel A, Geneser F", "2015", "Editorial Médica Panamericana", ""],
    ["", "Histología : atlas en color y texto", "Gartner LP", "2018 (7a ed.)", "Wolters Kluwer", "También 7a ed. (2018) en inglés"],
    ["", "Texto de histología : atlas a color", "Gartner LP", "2017 (7a ed.)", "Elsevier", "También 4a ed. (2017) en inglés"],
    ["", "Histología básica : texto y atlas", "Junqueira LCU, Carneiro J", "2015 (12a ed.)", "Editorial Médica Panamericana", "También 14a ed. en inglés"],
    ["", "Histology and cell biology : an introduction to pathology", "Kierszenbaum AL, Tres LL", "2016 (4th ed.)", "Elsevier/Saunders", "También 4a ed. (2016) en castellano"],
    ["", "Atlas color de citología e histología", "Kühnel W", "2005 (11a ed. corr. y)", "Medica Panamericana", "También 4th ed. (2003) en inglés."],
    ["", "Stevens y Lowe histología humana: cuarta edición", "Lowe JS, Anderson PG", "2015", "Elsevier", "También 4a ed. en inglés"],
    ["", "Histology: a text and atlas : with correlated cell and molecular biology", "Ross MH, Pawlina W", "2016 (7th ed.)", "Wolters Kluwer Health", "También 7a ed. en castellano"],
    ["", "Wheater : histología funcional : texto y atlas en color", "Young B, O’Dowd G, Woodford P", "2014 (6a ed.)", "Elsevier", "También 6a ed. en inglés"],
    ["", "http://www.udel.edu/biology/Wags/histopage/histopage.htm", "", "", "", "Pàgina web"],
    ["", "http://www.kumc.edu/instruction/medicine/anatomy/histoweb/index.htm", "", "", "", "Pàgina web"],
    ["", "http://www.lumen.luc.edu/lumen/MedEd/Histo/frames/histo_frames.html", "", "", "", "Pàgina web"],
    ["", "http://cal.vet.upenn.edu/index.php?page=histology", "", "", "", "Pàgina web"],
    ["", "https://ubmedicina.ventana-vector.com/", "", "", "", "Pàgina web"],
]

for fila in histologia:
    ws.append(fila)

# Introducción a la Salud, Antropología, Demografía. Historia de la Medicina
intro_salud = [
    ["Introducción a la Salud, Antropología, Demografía. Historia de la Medicina", "¿Me está escuchando doctor? : un viaje por la mente de los médicos", "Groopman JE", "2008", "RBA", ""],
    ["", "Breve historia de la medicina", "López Piñero JM", "2000", "Alianza", ""],
    ["", "Semmelweis", "Louis-Ferdinand Celine", "2014", "Ed Marbot", ""]
]

for fila in intro_salud:
    ws.append(fila)

# Inglés Médico
ws.append([
    "Inglés Médico", "Text electrònic — Articles actuals rellevants i recursos multimèdia — Materials publicats al Campus Virtual", "", "", "", ""
])

# Guardar archivo
output_file = "Asignaturas_Medicina_Completo.xlsx"
wb.save(output_file)
print("Archivo generado como Asignaturas_Medicina_Completo.xlsx")
