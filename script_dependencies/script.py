import sys
import time
import pdfplumber
import pandas as pd
import os
import datetime
import pytesseract # Importation pour l'OCR
from PIL import Image # Importation pour la manipulation d'images (pour l'OCR)
import fitz # PyMuPDF for converting PDF page to image (page.to_image().original is better with PyMuPDF)
import re # Importation du module re pour les expressions régulières avancées
import openpyxl # Importation pour manipuler les fichiers Excel existants
import numpy as np # Importation de NumPy pour une meilleure gestion de NaN  

log_dir = os.path.join(os.path.dirname(__file__), "logs")
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, f"log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
sys.stdout = open(log_file, 'w', encoding='utf-8')
sys.stderr = sys.stdout


    
def run_extraction(pdf_path):
    
    # IMPORTANT: Ensure Tesseract OCR engine is installed on your system
    # For macOS with Homebrew: brew brew install tesseract
    # For Windows: Download installer from https://tesseract-ocr.github.io/tessdoc/Downloads.html
    # For Linux: sudo apt-get install tesseract-ocr
    # Also ensure pytesseract Python library is installed: pip install pytesseract
    # And PyMuPDF: pip install PyMuPDF
    # Pour openpyxl: pip install openpyxl
    # Pour numpy: pip install numpy
    
    # --- PART 0 : Configuration Tesseract OCR (si utilisé) ---
    # Si tesseract n'est pas dans votre PATH système, vous devrez spécifier son chemin ici.
    # DÉCOMMENTEZ ET AJUSTEZ SI NÉCESSAIRE.
    pytesseract.pytesseract.tesseract_cmd = r'/opt/homebrew/bin/tesseract' # Exemple pour Mac M1/M2/M3
    
    # --- PART 1 : Configuration et Chemins ---
    
    # Nom de votre fichier PDF à traiter.
    # Assurez-vous que ce PDF se trouve dans le MÊME dossier que ce script Python.
    # <<< CONFIGURATION : Réglez ceci sur le PDF que vous traitez
    
    # Construire le chemin complet vers le PDF
    pdf_file_path = os.path.join(os.path.dirname(__file__), pdf_path)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0].lower()
    
    # Nom du fichier Excel de sortie détaillé (par page)
    # Il sera créé dans le MÊME dossier que votre script, avec un horodatage unique.
    timestamp = datetime.datetime.now().strftime("%d-%m-%Y")
    output_excel_detailed_file = os.path.join(os.path.dirname(__file__), f'Extration_{base_name}_{timestamp}.xlsx')
    
    # NOUVEAU: Chemin vers votre fichier Excel de MODÈLE (le template à remplir)
    # Assurez-vous que ce fichier TEMPLATE se trouve dans le MÊME dossier que ce script Python.
    template_excel_file_name = 'TEMPLATE.xlsx' # <<< À ADAPTER: Nom EXACT de votre fichier modèle Excel (ex: 'MonBilanTemplate.xlsx')
    template_excel_file_path = os.path.join(os.path.dirname(__file__), template_excel_file_name)
    
    # NOUVEAU: Nom du fichier Excel de sortie du MODÈLE REMPLI
    filled_template_output_file = os.path.join(os.path.dirname(__file__), f'Modèle_{base_name}_{timestamp}.xlsx')
    
    # NOUVEAU: Chemin vers votre fichier Excel de configuration des métriques
    # Ce fichier contiendra toutes vos métriques, libellés de recherche, feuilles sources/cibles, etc.
    metrics_config_file_name = 'Config_metrics.xlsx' # <<< NOUVEAU: Nom de votre fichier de configuration
    metrics_config_file_path = os.path.join(os.path.dirname(__file__), metrics_config_file_name)
    
    # --- PART 1.5: Chargement de la configuration des métriques depuis le fichier Excel ---
    financial_metrics_to_extract = {} # Initialiser le dictionnaire vide
    try:
        if os.path.exists(metrics_config_file_path):
            print(f"\nChargement de la configuration des métriques depuis '{metrics_config_file_name}'...")
            df_metrics_config = pd.read_excel(metrics_config_file_path)
            
            # --- DEBUG: Afficher les colonnes lues depuis Config_metrics.xlsx ---
            print(f"DEBUG: Colonnes lues de Config_metrics.xlsx: {df_metrics_config.columns.tolist()}")
            
            # Parcourir chaque ligne du DataFrame pour construire le dictionnaire
            for index, row in df_metrics_config.iterrows():
                # Nettoyage des espaces de début/fin (y compris les espaces insécables)
                metric_base_name = str(row['Metric_Base_Name']).strip()
                search_label = str(row['Search_Label']).strip()
                source_sheet = str(row['Source_Sheet_Name']).strip()
                target_cell = str(row['Target_Cell']).strip() # Ensure target_cell is also stripped
    
                raw_offset_col = row['Offset_Col']
                offset_col_str = str(raw_offset_col).strip()
                
                try:
                    offset_col_int = int(offset_col_str)
                except ValueError:
                    print(f"ERREUR: Impossible de convertir '{offset_col_str}' en nombre entier pour Offset_Col. Cette ligne sera ignorée. Métrique: '{metric_base_name}'")
                    continue
    
                # CRÉATION D'UNE CLÉ UNIQUE pour le dictionnaire financial_metrics_to_extract
                unique_metric_key = f"{metric_base_name} - {source_sheet} - Offset_{offset_col_int}"
                
                financial_metrics_to_extract[unique_metric_key] = {
                    'libellé_recherche': search_label,
                    'source_sheet_name': source_sheet,
                    'target_sheet_name': row['Target_Sheet_Name'],
                    'target_cell': target_cell, # Use the stripped target_cell
                    'offset_col': offset_col_int
                }
            print(f"Configuration des métriques chargée avec {len(financial_metrics_to_extract)} entrées.")
            
            print("\nDEBUG: Toutes les clés de métriques chargées:")
            for key in financial_metrics_to_extract.keys():
                print(f"  - {key}")
            print("-" * 50)
    
        else:
            print(f"ERREUR: Le fichier de configuration '{metrics_config_file_name}' n'a pas été trouvé à l'emplacement : {metrics_config_file_path}")
            print("Veuillez créer ce fichier ou vérifier son nom/chemin.")
    except Exception as e:
        print(f"ERREUR lors du chargement du fichier de configuration des métriques : {e}")
    
    
    # Dictionnaire pour stocker les données financières spécifiques extraites (clé_config_unique: valeur_numérique)
    extracted_financial_data = {}
    
    # --- Fonction utilitaire pour normaliser le texte pour la comparaison (Supprime tous les non-alphanumériques pour une recherche robuste) ---
    def normalize_text_for_comparison(text):
        if pd.isna(text):
            return ""
        text = str(text).strip()
        text = re.sub(r'\(cid:.*?\)', '', text) # Supprime les caractères spéciaux (cid:...)
        text = re.sub(r'[^a-zA-Z0-9]', '', text) # Supprime TOUS les caractères non alphanumériques (espaces, ponctuation, etc.)
        return text.lower() # Convertit en minuscules
    
    # --- PART 2 : Extraction, Nettoyage et Exportation des Tableaux ---
    
    print(f"Préparation de l'exportation des données par page vers '{output_excel_detailed_file}'...")
    writer = pd.ExcelWriter(output_excel_detailed_file, engine='xlsxwriter')
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            print(f"PDF '{pdf_path}' ouvert avec succès. Nombre de pages: {len(pdf.pages)}")
    
            for page_num, page in enumerate(pdf.pages, start=1):
                # --- CORRECTION ICI : Définir sheet_name au début de la boucle ---
                sheet_name = f'Page_{page_num}'
                # --- FIN DE LA CORRECTION ---
    
                print(f"\n--- Traitement de la Page {page_num} ---")
    
                df_page_extracted = pd.DataFrame()
                extraction_method = "Aucune"
    
                print("Tentative d'extraction avec PDFplumber (stratégie 'lines')...")
                table_settings_pdfplumber = {
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "snap_tolerance": 3,
                    "text_tolerance": 1,
                    "join_tolerance": 3,
                    "edge_min_length": 3,
                }
                try:
                    tables_pdfplumber = page.extract_tables(table_settings=table_settings_pdfplumber)
                    valid_tables_pdfplumber = [t for t in tables_pdfplumber if t and len(t) > 0 and any(any(cell and str(cell).strip() for cell in row) for row in t)]
                except Exception as e:
                    print(f"Erreur lors de l'extraction 'lines' avec PDFplumber: {e}")
                    valid_tables_pdfplumber = []
    
                if valid_tables_pdfplumber:
                    print("Tables significatives trouvées avec PDFplumber 'lines'.")
                    dfs_concat_temp = []
                    for table_data in valid_tables_pdfplumber:
                        df_table = pd.DataFrame(table_data)
                        dfs_concat_temp.append(df_table)
                    
                    if dfs_concat_temp:
                        df_page_extracted = pd.concat(dfs_concat_temp, ignore_index=True)
                        extraction_method = "PDFplumber (lignes)"
                    else:
                        print("Aucun DataFrame valide de PDFplumber après concaténation. Fallback à OCR.")
    
                if df_page_extracted.empty: # Si PDFplumber n'a rien trouvé ou que le DataFrame est vide
                    # --- TENTATIVE 2: Fallback à OCR si PDFplumber ne détecte rien ---
                    print("PDFplumber n'a pas détecté de tableaux significatifs ou le DataFrame est vide.")
                    print(f"Tentative d'extraction via OCR (Tesseract) pour la page {page_num}...")
                    extraction_method = "OCR (Tesseract)" 
                    
                    try:
                        # Convertir la page PDF en image Pillow pour Tesseract
                        pil_image = page.to_image(resolution=300).original # Higher resolution for better OCR
    
                        # Configuration OCR pour Tesseract - ESSAI DE PSM DIVERS
                        # --psm 3: Default, fully automatic page segmentation.
                        # --psm 6: Assume a single uniform block of text. (Souvent bon pour les tables)
                        # --psm 11: Sparse text. Find as much text as possible in no particular order.
                        ocr_config = r'--psm 6' # <<< CHANGEMENT CLÉ ICI
    
                        # Utiliser Tesseract pour extraire le texte avec les boîtes de délimitation
                        ocr_data = pytesseract.image_to_data(pil_image, output_type=pytesseract.Output.DATAFRAME, lang='fra', config=ocr_config)
                        
                        # Nettoyage des lignes vides et NaN du DataFrame OCR
                        ocr_data.dropna(subset=['text'], inplace=True)
                        ocr_data = ocr_data[ocr_data['text'].str.strip() != '']
                        
                        extracted_rows_ocr = []
                        
                        # Tri par position pour reconstruire les lignes
                        ocr_data_sod = ocr_data.sort_values(by=['page_num', 'block_num', 'par_num', 'line_num', 'left'])
                        
                        current_line_y_tolerance = 10 # Tolérance pour considérer les mots comme étant sur la même ligne Y
                        current_col_x_tolerance = 30 # Tolérance pour grouper les mots dans la même colonne X
                        
                        line_groups = {}
                        for idx, row in ocr_data_sod.iterrows():
                            key = (row['page_num'], row['block_num'], row['par_num'], row['line_num'])
                            if key not in line_groups:
                                line_groups[key] = []
                            line_groups[key].append(row)
                        
                        for key in sod(line_groups.keys()):
                            words_in_line = line_groups[key]
                            
                            columns_in_line = []
                            
                            for word_data in words_in_line:
                                word_text = str(word_data['text']).strip()
                                word_left = word_data['left']
                                
                                assigned_to_col = False
                                for col_idx, col_entry in enumerate(columns_in_line):
                                    if abs(word_left - col_entry['left_min']) < current_col_x_tolerance or \
                                       (word_left >= col_entry['left_min'] and word_left <= col_entry['left_max'] + current_col_x_tolerance):
                                        
                                        columns_in_line[col_idx]['text'] += " " + word_text
                                        columns_in_line[col_idx]['left_max'] = max(columns_in_line[col_idx]['left_max'], word_data['left'] + word_data['width'])
                                        assigned_to_col = True
                                        break
                                
                                if not assigned_to_col:
                                    columns_in_line.append({
                                        'text': word_text,
                                        'left_min': word_left,
                                        'left_max': word_left + word_data['width']
                                    })
                            
                            # Extraire le texte de chaque colonne et les trier par position X
                            sod_cells = [col['text'].strip() for col in sod(columns_in_line, key=lambda x: x['left_min'])]
                            extracted_rows_ocr.append(sod_cells)
                            
                        if extracted_rows_ocr:
                            df_page_extracted = pd.DataFrame(extracted_rows_ocr)
                            print(f"OCR a extrait un DataFrame avec {len(df_page_extracted)} lignes pour la page {page_num}.")
                            print("ATTENTION: L'extraction OCR est moins précise et peut nécessiter un nettoyage manuel des données et des en-têtes.")
                        else:
                            print(f"OCR n'a pas pu extraire de texte structuré en tableau pour la page {page_num}.")
                            df_page_extracted = pd.DataFrame() 
                    
                    except pytesseract.TesseractNotFoundError:
                        print("ERREUR Tesseract: Le moteur Tesseract OCR n'est pas installé ou n'est pas dans votre PATH.")
                        print("Veuillez l'installer et/ou configurer pytesseract.pytesseract.tesseract_cmd si nécessaire.")
                        df_page_extracted = pd.DataFrame()
                    except Exception as e:
                        print(f"ERREUR lors du traitement OCR pour la page {page_num} : {e}")
                        df_page_extracted = pd.DataFrame()
    
    
                # --- Aperçu du DataFrame après les tentatives d'extraction (avant nettoyage final) ---
                if not df_page_extracted.empty:
                    print(f"\n--- Aperçu du DataFrame extrait (avant nettoyage final) Page {page_num} ({extraction_method}) ---")
                    print(df_page_extracted.head())
                    print(f"Le DataFrame extrait de la page {page_num} a {len(df_page_extracted)} lignes (avant nettoyage).")
                    print("-----------------------------------------------------")
                else:
                    print(f"\n--- Aucun DataFrame significatif extrait pour la page {page_num}. ---")
    
    
                # --- Nettoyage des Données pour cette page ---
                if not df_page_extracted.empty:
                    # Définition de header_rows_to_skip en fonction du nom du fichier
                    # Cela permet d'adapter la gestion des en-têtes par document.
                    header_rows_to_skip = 1 # Valeur par défaut si aucun match
                    print(f"Aucune configuration spécifique pour '{pdf_path}'. `header_rows_to_skip` reste à {header_rows_to_skip}.")
    
    
                    # Logique pour définir les en-têtes et sauter les lignes.
                    if header_rows_to_skip > 0 and len(df_page_extracted) >= header_rows_to_skip:
                        if header_rows_to_skip == 2: # Cas ETE/ (combinaison de 2 lignes)
                            combined_headers = (df_page_extracted.iloc[0].astype(str).fillna('') + " " + \
                                                df_page_extracted.iloc[1].astype(str).fillna('')).str.strip()
                            # Rendre les en-têtes uniques pour éviter les erreurs de Pandas
                            unique_headers = []
                            seen = {}
                            for h in combined_headers:
                                temp_h = h
                                counter = 1
                                while temp_h in seen:
                                    temp_h = f"{h}_{counter}"
                                    counter += 1
                                unique_headers.append(temp_h)
                                seen[temp_h] = True
                            df_page_extracted.columns = unique_headers
                            print(f"En-têtes combinées pour la page {page_num} : {unique_headers}")
                            df_page_extracted = df_page_extracted[header_rows_to_skip:].reset_index(drop=True)
                        elif header_rows_to_skip == 1: # Cas  (première ligne seule)
                            df_page_extracted.columns = df_page_extracted.iloc[0].astype(str)
                            print(f"En-têtes définies pour la page {page_num} à partir de la ligne 0 : {df_page_extracted.columns.tolist()}")
                            df_page_extracted = df_page_extracted[header_rows_to_skip:].reset_index(drop=True)
                        else: # Si header_rows_to_skip est une autre valeur et non gérée spécifiquement
                            print(f"Attention: `header_rows_to_skip` est défini sur {header_rows_to_skip} mais la logique de combinaison/saut de lignes n'est pas adaptée ou il n'y a pas assez de lignes.")
                            df_page_extracted = df_page_extracted[header_rows_to_skip:].reset_index(drop=True)
                    else:
                        print(f"Aucune ligne d'en-tête n'est sautée pour la page {page_num}. Les colonnes numériques par défaut sont utilisées.")
                    
                    # --- Débogage conversion numérique Page {page_num} ---
                    # Affiche les 5 premières valeurs de chaque colonne après nettoyage initial (avant to_numeric)
                    print(f"\n--- Débogage conversion numérique Page {page_num} ---")
                    for col in df_page_extracted.columns.tolist():
                        original_stripped_series = df_page_extracted[col].astype(str).str.strip()
                        # Supprime (cid: n'impo quoi) puis les espaces normaux et remplace la virgule.
                        # Remplacement de tous les espaces par un seul pour l'affichage de débogage
                        numeric_candidate = original_stripped_series.str.replace(r'\(cid:.*?\)', '', regex=True) \
                                                                    .str.replace(r'\s+', ' ', regex=True) 
                        print(f"  Colonne '{col}': Premières valeurs brutes nettoyées pour conversion:")
                        print(numeric_candidate.head(5)) # Afficher les 5 premières valeurs
                    print("--------------------------------------------------")
    
                    # Détection dynamique des types de colonnes (texte ou numérique)
                    column_types = {}
                    # NOUVEAU: Liste de patterns pour les colonnes qui sont SENSÉES être numériques pour 
                    # Ceci aidera à forcer la détection numérique pour ces colonnes.
                    expected_numeric_patterns_for_ = [
                        'exercice', 'précédent', 'brut', 'net', 'amortissements', 'provisions', 'opérations propres',
                        'opérations concernant', 'totaux de l\'exercice', 'totaux de l\'exercice précédent'
                    ]
    
                    # Ajout de la vérification ici pour éviter l'erreur "truth value of a Series is ambiguous"
                    if not df_page_extracted.empty: 
                        print(f"DEBUG: DataFrame non vide pour la page {page_num} avant la détection des types de colonnes.")
                        for col in df_page_extracted.columns.tolist():
                            # Initialisation pour la détection
                            is_expected_numeric_by_name = False
                            # Vérifier si le nom de la colonne correspond à un motif numérique attendu
                            for pattern in expected_numeric_patterns_for_:
                                if pattern in str(col).lower():
                                    is_expected_numeric_by_name = True
                                    break
                            
                            # Nettoyage pour l'évaluation du type (retire cid, tous les espaces, remplace virgule, *et* caractères non-numériques)
                            # Ceci est la CLÉ pour une détection robuste des colonnes numériques
                            temp_cleaned_for_type_check = df_page_extracted[col].astype(str).str.strip() \
                                                                                  .str.replace(r'\(cid:.*?\)', '', regex=True) \
                                                                                  .str.replace(r'\s+', '', regex=True) \
                                                                                  .str.replace(',', '.', regex=False)
                            # Correction: Supprimer tout sauf les chiffres, points et tirets pour la DÉTECTION
                            temp_cleaned_for_type_check = temp_cleaned_for_type_check.apply(
                                lambda x: re.sub(r'[^\d\.\-]+', '', str(x)) if pd.notna(x) else x
                            )
                            
                            # DEBUG: Affiche les valeurs après ce nettoyage agressif pour la détection
                            print(f"  Détection type colonne '{col}': Valeurs après nettoyage agressif pour test:")
                            print(temp_cleaned_for_type_check.head(5))
    
    
                            # Tentative de conversion pour compter les succès
                            numeric_count = pd.to_numeric(temp_cleaned_for_type_check, errors='coerce').count()
                            total_count = temp_cleaned_for_type_check.count() # Compte les non-NaN après le nettoyage
                            
                            if total_count == 0: # Si la colonne est entièrement vide
                                column_types[col] = 'text' # Traitez-la comme du texte pour éviter des erreurs plus tard
                                print(f"  (Col '{col}' - Vide, classée texte)")
                                continue
    
                            # Seuil: Si plus de 70% des cellules non-vides peuvent être converties en nombre, c'est numérique.
                            # OU SI le nom de la colonne est dans les patterns numériques attendus.
                            if (numeric_count / total_count > 0.70) or is_expected_numeric_by_name: 
                                column_types[col] = 'numeric'
                            else:
                                column_types[col] = 'text'
                            
                            print(f"  (Col '{col}' - numeric_count: {numeric_count}, total_count: {total_count}, ratio: {numeric_count/total_count:.2f}, expected_numeric: {is_expected_numeric_by_name}, classée: {column_types[col]})")
                        
                        print(f"Détection dynamique finale des types de colonnes pour la page {page_num}: {column_types}")
    
                        # Nettoyage et conversion des colonnes basé sur la détection dynamique
                        for col in df_page_extracted.columns.tolist():
                            current_series_as_str = df_page_extracted[col].astype(str).str.strip()
    
                            # Étape 1: Nettoyage universel de (cid:...)
                            cleaned_series_base = current_series_as_str.str.replace(r'\(cid:.*?\)', '', regex=True)
                            
                            if column_types[col] == 'numeric':
                                # Pour les colonnes numériques : supprimer TOUS les espaces, remplacer virgule par point
                                numeric_candidate_str = cleaned_series_base.str.replace(r'\s+', '', regex=True) \
                                                                               .str.replace(',', '.', regex=False)
                                
                                # Correction: Appliquer le nettoyage agressif aussi à la conversion finale
                                final_numeric_string_for_conversion = numeric_candidate_str.apply(
                                    lambda x: re.sub(r'[^\d\.\-]+', '', str(x)) if pd.notna(x) else x
                                )
    
                                conved_numeric_series = pd.to_numeric(final_numeric_string_for_conversion, errors='coerce')
                                df_page_extracted[col] = conved_numeric_series # Les NaN restent à ce stade
                                print(f"  Colonne '{col}': Traitée comme colonne numérique (nettoyage agressif appliqué).")
                            else: # Traiter comme texte
                                # Pour les colonnes de texte : remplacer plusieurs espaces par un seul espace (conserver le formatage textuel)
                                df_page_extracted[col] = cleaned_series_base.str.replace(r'\s+', ' ', regex=True).fillna('')
                                print(f"  Colonne '{col}': Traitée comme colonne de texte (pas de conversion numérique).")
                                
                        # Remplir les NaN après le traitement des colonnes
                        # Les colonnes numériques (détectées comme telles) reçoivent 0 pour les NaN
                        # Les colonnes de texte (détectées comme telles) reçoivent une chaîne vide pour les NaN
                        for col in df_page_extracted.columns.tolist():
                            if column_types[col] == 'numeric':
                                df_page_extracted[col] = df_page_extracted[col].fillna(0)
                            else:
                                df_page_extracted[col] = df_page_extracted[col].fillna('')
    
                    else: # Si df_page_extracted est vide au début de la détection de type
                        print(f"ATTENTION: Le DataFrame de la page {page_num} est vide avant la détection des types de colonnes. Saut des étapes de nettoyage et de conversion.")
                        continue # Passe à la page suivante si le DataFrame est vide
    
    
                    # Supprimer les lignes entièrement vides (où toutes les valeurs sont NaN ou chaînes vides)
                    df_page_extracted.dropna(how='all', inplace=True)
                    
                    # Supprimer les colonnes où TOUTES les valeurs sont NaN ou chaînes vides (colonnes vides)
                    df_page_extracted.dropna(axis=1, how='all', inplace=True)
                    
                    # Re-indexer après le dropna
                    df_page_extracted.reset_index(drop=True, inplace=True)
    
                    print(f"Le DataFrame final de la page {page_num} a {len(df_page_extracted)} lignes après nettoyage.")
                    print("\n--- Aperçu du DataFrame final (après nettoyage et conversion) ---")
                    print(df_page_extracted.head())
                    print(f"Types de données du DataFrame final de la page {page_num} :")
                    print(df_page_extracted.dtypes) # Diagnostic des types de données
                    print("------------------------------------------------------------")
    
    
                    # --- NOUVEAUX DÉBOGAGES JUSTE AVANT L'EXPORTATION DU FICHIER DÉTAILLÉ ---
                    print(f"\nDEBUG EXCEL_DETAIL_EXPORT: Préparation de l'exportation de la Page {page_num} vers la feuille '{sheet_name}'.")
                    print("DEBUG EXCEL_DETAIL_EXPORT: Shape of DataFrame:", df_page_extracted.shape)
                    print("DEBUG EXCEL_DETAIL_EXPORT: Columns:", df_page_extracted.columns.tolist())
                    print("DEBUG EXCEL_DETAIL_EXPORT: Dtypes:\n", df_page_extracted.dtypes)
                    print("DEBUG EXCEL_DETAIL_EXPORT: First 5 rows:\n", df_page_extracted.head().to_string())
                    print("DEBUG EXCEL_DETAIL_EXPORT: Has any NaN in columns?", df_page_extracted.isnull().any())
                    # --- FIN DES NOUVEAUX DÉBOGAGES ---
    
                    # --- NOUVEAU: Renommer les colonnes vides ou 'None' pour une meilleure compatibilité Excel ---
                    new_columns = []
                    for i, col in enumerate(df_page_extracted.columns):
                        # Correction: utiliser str(col) pour gérer les cas où col est None ou np.nan
                        col_str = str(col).strip().lower()
                        if pd.isna(col) or col_str == '' or col_str == 'none':
                            new_columns.append(f'Unnamed_Col_{i}')
                            print(f"  Renommage colonne: '{col}' (index {i}) -> 'Unnamed_Col_{i}'")
                        else:
                            new_columns.append(col)
                    df_page_extracted.columns = new_columns
                    print(f"DEBUG EXCEL_DETAIL_EXPORT: Colonnes après renommage: {df_page_extracted.columns.tolist()}")
                    # --- FIN NOUVEAU RENOMMAGE ---
    
                    # --- Exportation vers une feuille Excel ---
                    print(f"Exportation de la Page {page_num} vers la feuille '{sheet_name}'...")
                    try:
                        df_page_extracted.to_excel(
                            writer,
                            sheet_name=sheet_name,
                            index=False,
                            header=True, # Conserve True pour que les noms de colonnes numériques (0, 1, 2...) soient exportés
                            float_format="%.2f"
                        )
                        print(f"Page {page_num} exportée avec succès.")
                    except Exception as e:
                        print(f"Une erreur est survenue lors de l'exportation de la page {page_num} vers Excel : {e}")
                else:
                    print(f"Le DataFrame de la page {page_num} est vide après nettoyage. Pas d'exportation.")
    
    except FileNotFoundError:
        print(f"ERREUR: Le fichier PDF '{pdf_path}' n'a pas été trouvé à l'emplacement : {pdf_file_path}")
        print("Veuillez vous assurer que le PDF est dans le même dossier que le script ou que le chemin est correct.")
    except Exception as e:
        print(f"Une erreur inattendue est survenue lors du traitement du PDF : {e}")
    
    finally: # Ce bloc garantit que le writer est toujours fermé.
        try:
            writer.close() # Utilisez close() pour les versions récentes de Pandas
            print(f"\nTous les tableaux ont été exports avec succès dans '{os.path.basename(output_excel_detailed_file)}'.")
            print(f"Le fichier Excel détaillé se trouve dans : {os.path.dirname(output_excel_detailed_file)}")
        except Exception as e:
            print(f"Une erreur est survenue lors de la finalisation du fichier Excel détaillé : {e}")
    
        detailed_excel_data = {}
        try:
            if os.path.exists(output_excel_detailed_file):
                print(f"\nChargement du fichier Excel détaillé '{os.path.basename(output_excel_detailed_file)}' pour la synthèse...")
                xls = pd.ExcelFile(output_excel_detailed_file)
                for sheet_name_loaded in xls.sheet_names: # Renommé pour éviter conflit avec sheet_name du scope parent
                    detailed_excel_data[sheet_name_loaded] = pd.read_excel(xls, sheet_name=sheet_name_loaded, header=0)
                print(f"Fichier détaillé chargé. Feuilles disponibles: {list(detailed_excel_data.keys())}")
            else:
                print(f"ERREUR: Le fichier Excel détaillé '{os.path.basename(output_excel_detailed_file)}' n'a pas été trouvé. Impossible de remplir le modèle.")
                detailed_excel_data = None
        except Exception as e:
            print(f"ERREUR lors du chargement du fichier Excel détaillé : {e}")
            detailed_excel_data = None
    
        # --- PARTIE FINALE : Extraction des Métriques Spécifiques et Remplissage du TEMPLATE Excel ---
        if detailed_excel_data and financial_metrics_to_extract:
            print(f"\n--- Remplissage du MODÈLE Excel '{template_excel_file_name}' ---")
            
            # POPULATION DU DICTIONNAIRE extracted_financial_data
            for label_config_key, config in financial_metrics_to_extract.items():
                libelle_recherche = config['libellé_recherche']
                source_sheet_name = config['source_sheet_name']
                offset_col = config['offset_col']
                target_cell = config['target_cell'] # Get target_cell here
    
                # --- NOUVELLE VALIDATION : Vérifier si Target_Cell est vide ---
                if not target_cell:
                    print(f"ERREUR: La cellule cible (Target_Cell) est vide pour la métrique '{label_config_key}'. Cette métrique sera ignorée. Veuillez remplir la colonne 'Target_Cell' dans votre fichier '{metrics_config_file_name}'.")
                    continue # Passer à la métrique suivante
                # --- FIN DE LA VALIDATION ---
    
    
                metric_value_found_and_extracted = False 
                
                if source_sheet_name in detailed_excel_data:
                    source_df = detailed_excel_data[source_sheet_name]
                    
                    # Normaliser le libellé de recherche pour la comparaison
                    normalized_search_label = normalize_text_for_comparison(libelle_recherche)
                    print(f"\nDEBUG Recherche: Libellé de recherche original: '{libelle_recherche}' -> Normalisé: '{normalized_search_label}'")
    
    
                    for col_name in source_df.columns:
                        # Normaliser le contenu de la colonne source pour la correspondance
                        normalized_source_column_for_match = source_df[col_name].apply(normalize_text_for_comparison)
                        
                        # DEBUG: Afficher les premières valeurs normalisées de la colonne source pour les libellés spécifiques
                        # Ajustez cette condition pour les libellés que vous souhaitez déboguer spécifiquement
                        if "banquestgetccp" in normalized_search_label or \
                           "titresvaleursdeplacementh" in normalized_search_label or \
                           "fraispreliminaires" in normalized_search_label: # Add more specific debug targets here if needed
                            print(f"  DEBUG Colonne '{col_name}': Premières 5 valeurs normalisées pour correspondance:")
                            print(normalized_source_column_for_match.head(5).to_string()) # Utiliser .to_string() pour un affichage propre
    
                        # Effectuer la recherche avec la chaîne normalisée (regex=False car déjà nettoyée)
                        matching_rows = source_df[
                            normalized_source_column_for_match == normalized_search_label
                        ]

                        # fallback to contains if nothing is found
                        if matching_rows.empty:
                            matching_rows = source_df[
                                normalized_source_column_for_match.str.contains(normalized_search_label, na=False, regex=False)
                            ]

                        
                        if not matching_rows.empty:
                            matching_row_series = matching_rows.iloc[0]
                            col_index_of_label = source_df.columns.get_loc(col_name)
                            value_col_index = col_index_of_label + offset_col
                            
                            if value_col_index < len(source_df.columns):
                                raw_extracted_value = matching_row_series[source_df.columns[value_col_index]] # Get value by column name from raw_extracted_value
    
                                # Tenter de convertir la valeur en numérique de manière robuste (améliorée pour formats européens)
                                final_value = np.nan # Initialiser à NaN
                                print(f"\n    --- DEBUG CONVERSION POUR '{label_config_key}' (Cellule source: '{source_df.columns[value_col_index]}') ---")
                                print(f"    DEBUG CONVERSION: Valeur brute extraite: '{raw_extracted_value}' (Type: {type(raw_extracted_value)})")
    
                                if pd.api.types.is_numeric_dtype(raw_extracted_value):
                                    # Si c'est déjà un type numérique (int ou float), utiliser directement
                                    final_value = float(raw_extracted_value)
                                    print(f"    DEBUG CONVERSION: La valeur est déjà numérique. Valeur finale: {final_value}")
                                elif pd.notna(raw_extracted_value): # Seulement si ce n'est pas NaN et pas numérique (donc probablement une chaîne)
                                    cleaned_val_str = str(raw_extracted_value).strip()
                                    print(f"    DEBUG CONVERSION: Valeur chaîne brute: '{cleaned_val_str}'")
                                    
                                    # 1. Supprime (cid:...)
                                    cleaned_val_str = re.sub(r'\(cid:.*?\)', '', cleaned_val_str)
                                    print(f"    DEBUG CONVERSION: Après suppression (cid:): '{cleaned_val_str}'")
                                    
                                    # 2. Supprimer tous les espaces (normaux et insécables)
                                    cleaned_val_str = cleaned_val_str.replace('\xa0', '').replace(' ', '')
                                    print(f"    DEBUG CONVERSION: Après suppression des espaces: '{cleaned_val_str}'")
    
                                    # 3. Gérer les séparateurs décimaux et de milliers
                                    # Compo à la fois un point et une virgule (ex: 1.234.567,89)
                                    if ',' in cleaned_val_str and '.' in cleaned_val_str:
                                        # Supprimer les points (séparateurs de milliers)
                                        cleaned_val_str = cleaned_val_str.replace('.', '')
                                        # Remplacer la virgule par un point décimal
                                        cleaned_val_str = cleaned_val_str.replace(',', '.')
                                        print(f"    DEBUG CONVERSION: Après gestion format Européen (points et virgule): '{cleaned_val_str}'")
                                    # Compo seulement une virgule (ex: 1234,56)
                                    elif ',' in cleaned_val_str:
                                        # Remplacer la virgule par un point décimal
                                        cleaned_val_str = cleaned_val_str.replace(',', '.')
                                        print(f"    DEBUG CONVERSION: Après gestion format Européen (virgule seulement): '{cleaned_val_str}'")
                                    # Si seul le point est présent (ex: "1234.56") ou aucun, pas de remplacement nécessaire
                                    
                                    # 4. Nettoyage final: Conserver uniquement chiffres, point décimal, et tiret (pour les négatifs)
                                    # Conserve explicitement les chiffres, le point décimal et le tiret.
                                    cleaned_val_str = re.sub(r'[^\d\.\-]+', '', cleaned_val_str)
                                    print(f"    DEBUG CONVERSION: Après nettoyage final regex: '{cleaned_val_str}'")
                                    
                                    # S'assurer qu'il n'y a qu'un seul point décimal et pas vide/juste un tiret
                                    if not cleaned_val_str or cleaned_val_str == '-' or cleaned_val_str == '.':
                                        final_value = np.nan
                                        print(f"    DEBUG CONVERSION: Chaîne vide/invalide après nettoyage, défini à NaN.")
                                    elif cleaned_val_str.count('.') > 1:
                                        final_value = np.nan # Ambigu, définir à NaN
                                        print(f"    DEBUG CONVERSION: Plusieurs points décimaux détectés, défini à NaN.")
                                    else:
                                        try:
                                            final_value = float(cleaned_val_str)
                                            print(f"    DEBUG CONVERSION: Converti en float avec succès: {final_value}")
                                        except ValueError:
                                            final_value = np.nan
                                            print(f"    DEBUG CONVERSION: Échec de conversion en float (ValueError), défini à NaN.")
                                else: # raw_extracted_value est NaN
                                    final_value = np.nan
                                    print(f"    DEBUG CONVERSION: Valeur brute est NaN, défini à NaN.")
                                
                                extracted_financial_data[label_config_key] = final_value
                                print(f"  Extrait de '{source_sheet_name}' (via '{libelle_recherche}'): '{label_config_key}' = {final_value} (type: {type(final_value)})")
                                metric_value_found_and_extracted = True
                                break # Valeur trouvée pour cette métrique, passer à la suivante
                            else:
                                print(f"  Avertissement: Offset de colonne ({offset_col}) hors des limites pour '{libelle_recherche}' sur feuille '{source_sheet_name}'.")
                                
                        if metric_value_found_and_extracted:
                            break # Libellé trouvé dans cette colonne, pas besoin de chercher dans d'autres colonnes de la même feuille
    
                if not metric_value_found_and_extracted:
                    print(f"  Avertissement: Libellé '{libelle_recherche}' non trouvé ou valeur non extraite sur feuille '{source_sheet_name}' du fichier détaillé. Défaut à NaN.")
                    extracted_financial_data[label_config_key] = np.nan # Assurer que NaN est mis si non trouvé
    
            print("\nDEBUG FINAL: Contenu de 'extracted_financial_data' avant écriture dans le modèle:")
            if extracted_financial_data:
                for key, value in extracted_financial_data.items():
                    print(f"  - Clé: '{key}', Valeur: '{value}'")
            else:
                print("  Le dictionnaire 'extracted_financial_data' est vide. Aucune valeur à écrire.")
            print("-" * 50)
    
            try:
                if not os.path.exists(template_excel_file_path):
                    print(f"ERREUR: Le fichier modèle Excel '{template_excel_file_name}' n'a pas été trouvé à l'emplacement : {template_excel_file_path}")
                    print("Veuillez vous assurer que le modèle est dans le même dossier que le script ou que le chemin est correct.")
                else:
                    workbook = openpyxl.load_workbook(template_excel_file_path)
    
                    for label_config_key, value in extracted_financial_data.items():
                        config = financial_metrics_to_extract.get(label_config_key)
                        if config:
                            target_sheet_name = config['target_sheet_name'] # Renommé pour éviter conflit avec sheet_name du scope parent
                            cell_address = config['target_cell']
    
                            # --- NOUVELLE VALIDATION : Vérifier si Target_Cell est vide juste avant d'écrire ---
                            if not cell_address:
                                print(f"ERREUR: La cellule cible (Target_Cell) est vide pour la métrique '{label_config_key}'. Écriture ignorée.")
                                continue # Passer à la métrique suivante
                            # --- FIN DE LA VALIDATION ---
    
                            try:
                                sheet = workbook[target_sheet_name]
                                
                                print(f"    DEBUG ÉCRITURE: Écriture de la valeur '{value}' (Type: {type(value)}) dans la cellule {cell_address} de la feuille '{target_sheet_name}'.")
                                # Logique d'écriture: Si NaN, laisser vide; si 0, écrire 0; sinon, écrire la valeur
                                if pd.isna(value):
                                    sheet[cell_address] = None # Laisse la cellule vide pour les NaN
                                    print(f"  Écrit '{label_config_key}' : Cellule vide (valeur non trouvée) dans la cellule {cell_address} de la feuille '{target_sheet_name}' du modèle.")
                                elif isinstance(value, (int, float)) and value == 0:
                                    sheet[cell_address] = 0 # Écrit le chiffre 0 pour les zéros numériques
                                    print(f"  Écrit '{label_config_key}' : 0 (zéro numérique) dans la cellule {cell_address} de la feuille '{target_sheet_name}' du modèle.")
                                else:
                                    sheet[cell_address] = value # Écrit la valeur telle quelle (numérique non nulle)
                                    print(f"  Écrit '{label_config_key}' : {value} dans la cellule {cell_address} de la feuille '{target_sheet_name}' du modèle.")
    
                            except KeyError:
                                print(f"  ERREUR: La feuille cible '{target_sheet_name}' n'existe pas dans le modèle Excel pour '{label_config_key}'. Vérifiez le nom de la feuille dans votre template.")
                            except Exception as e:
                                print(f"  Erreur lors de l'écriture de '{label_config_key}' dans la cellule {cell_address} de la feuille '{target_sheet_name}' du modèle : {e}")
                        else:
                            print(f"  Avertissement: Configuration non trouvée pour '{label_config_key}' lors de l'écriture dans le modèle.")
                    
                    # --- NOUVEAU: Try-except pour la sauvegarde du workbook ---
                    try:
                        workbook.save(filled_template_output_file)
                        print(f"Le modèle Excel a été rempli et enregistré sous '{os.path.basename(filled_template_output_file)}'.")
                        workbook.close()
                    except Exception as save_e:
                        print(f"ERREUR lors de la sauvegarde du modèle Excel rempli '{os.path.basename(filled_template_output_file)}' : {save_e}")
                        print("Vérifiez que le fichier n'est pas ouvert par Excel ou d'autres applications, et que vous avez les permissions d'écriture.")
                
            except Exception as e:
                print(f"Une erreur est survenue lors du remplissage du modèle Excel : {e}")
        else:
            print("\nAucune configuration de métriques n'a été chargée ou le fichier Excel détaillé est vide. Impossible de remplir le modèle.")
    
        print(f"\nProcessus d'extraction et d'exportation terminé.")
        print(f"Les fichiers de sortie se trouvent dans : {os.path.dirname(pdf_file_path)}")
    writer.close()
    xls.close()
    workbook.close()
    time.sleep(2)
