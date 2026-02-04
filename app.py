import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
from io import BytesIO
import zipfile
import re
from datetime import datetime
import copy

# --- KONFIGURACJA STRONY I DESIGN ---
st.set_page_config(page_title="Generator Katalogu CC", layout="wide", page_icon="📄")

# CSS - Style i Wygląd (Poprawiony kontrast i layout)
st.markdown("""
    <style>
        /* Globalny reset do bieli i czerni */
        .stApp {
            background-color: #FFFFFF;
            color: #000000;
        }
        /* Wymuszenie czarnego tekstu wszędzie */
        h1, h2, h3, h4, h5, h6, p, span, div, label, .stMarkdown {
            color: #000000 !important;
            font-family: 'Roboto', sans-serif !important;
        }
        /* Poprawa kontrastu w SideBarze (Uploader plików) */
        [data-testid="stSidebar"] {
            background-color: #F4F4F4; /* Jasnoszary dla odcięcia */
            border-right: 1px solid #000000;
        }
        [data-testid="stSidebar"] label {
            font-weight: bold;
            color: #000000 !important;
        }
        .stFileUploader div {
            color: #000000 !important;
        }
        .stFileUploader small {
            color: #333333 !important;
        }
        
        /* Przyciski */
        .stButton>button {
            width: 100%;
            border-radius: 0px;
            border: 2px solid #000000;
            height: 3em;
            background-color: #FFFFFF;
            color: #000000;
            font-weight: bold;
            transition: all 0.3s;
        }
        .stButton>button:hover {
            background-color: #000000;
            color: #FFFFFF;
            border-color: #000000;
        }
        
        /* Tabela */
        [data-testid="stDataFrame"] {
            border: 1px solid #000000;
        }
        
        /* Ukrycie zbędnych elementów */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# --- FUNKCJE LOGICZNE ---

def parse_scale_business(value):
    """Zamienia tekst '1 mln PLN' na liczbę, żeby dało się sortować."""
    if pd.isna(value): return 0.0
    text = str(value).lower().replace(',', '.').replace(' ', '')
    multiplier = 1.0
    if 'mld' in text or 'b' in text: multiplier = 1_000_000_000.0
    elif 'mln' in text or 'm' in text: multiplier = 1_000_000.0
    elif 'tys' in text or 'k' in text: multiplier = 1_000.0
    numbers = re.findall(r"[-+]?\d*\.\d+|\d+", text)
    if numbers: return float(numbers[0]) * multiplier
    return 0.0

def safe_duplicate_slide(pres, index):
    """
    Bezpieczniejsza wersja duplikowania slajdu.
    Zamiast insert_before (które powoduje błędy XML), używamy append.
    """
    try:
        source = pres.slides[index]
        blank_slide_layout = pres.slide_layouts[6] # Pusty layout
        dest = pres.slides.add_slide(blank_slide_layout)

        # Kopiowanie kształtów
        for shape in source.shapes:
            try:
                new_el = copy.deepcopy(shape.element)
                dest.shapes._spTree.append(new_el) # Używamy append zamiast insert_before
            except Exception:
                # Jeśli jakiś specyficzny kształt powoduje błąd, pomijamy go, ale slajd powstaje
                continue

        # Próba skopiowania relacji (np. tła, stylów), ale bez crashowania apki
        try:
            for key, value in source.part.rels.items():
                if "notesSlide" not in value.reltype:
                    dest.part.rels.add_relationship(value.reltype, value._target, value.rId)
        except:
            pass
            
        return dest
    except Exception as e:
        print(f"Critical Slide Error: {e}")
        return None

def clean_polish_typography(text):
    if not isinstance(text, str): return text
    conjunctions = [" w ", " z ", " i ", " a ", " o ", " u ", " na ", " do "]
    for word in conjunctions:
        pattern = re.compile(re.escape(word), re.IGNORECASE)
        text = pattern.sub(lambda m: m.group(0).replace(' ', '\u00A0', 1), text)
    return text

def replace_text_in_shape(shape, replacements):
    if not shape.has_text_frame: return
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            # Zachowujemy oryginalny tekst, żeby podmieniać w nim klucze
            original_text = run.text
            for key, value in replacements.items():
                if key in original_text:
                    new_val = str(value)
                    
                    # Logika typografii i wielkości czcionki tylko dla opisu
                    if key == "{Katalog Członków CC - opis do 500 znaków}":
                        new_val = clean_polish_typography(new_val)
                        if len(new_val) > 600: run.font.size = Pt(8)
                        elif len(new_val) > 450: run.font.size = Pt(9)
                    
                    original_text = original_text.replace(key, new_val)
            
            # Przypisujemy zmieniony tekst
            run.text = original_text

def replace_image_in_shape(slide, shape, image_stream):
    try:
        left, top = shape.left, shape.top
        width, height = shape.width, shape.height
        slide.shapes.add_picture(image_stream, left, top, width, height)
        # Usuwamy stary kształt (placeholder)
        sp = shape._element
        sp.getparent().remove(sp)
    except:
        pass

# --- INTERFEJS ---

st.title("Generator Katalogu CC")

# 1. SIDEBAR
with st.sidebar:
    st.header("1. Wgraj pliki")
    uploaded_excel = st.file_uploader("Baza Danych (Excel)", type=['xlsx', 'csv'])
    uploaded_pptx = st.file_uploader("Szablon (.pptx)", type=['pptx'])
    uploaded_zip = st.file_uploader("Zdjęcia (.zip)", type=['zip'])
    st.markdown("---")
    st.write("Wskazówka: Nazwy plików zdjęć w Excelu muszą pasować do plików w ZIP.")

# 2. GŁÓWNA LOGIKA
if uploaded_excel and uploaded_pptx:
    # Ładowanie danych
    if uploaded_excel.name.endswith('.csv'):
        df = pd.read_csv(uploaded_excel)
    else:
        df = pd.read_excel(uploaded_excel)
    
    df.columns = df.columns.str.strip()

    # Mapowanie kolumn
    wanted_cols = ["Imię", "Nazwisko", "Firma", "Branża", "Katalog Członków CC - opis do 500 znaków", "Grupa CC", "Skala Biznesu", "Photo", "Logo"]
    
    # Inteligentne szukanie nazw kolumn
    col_map = {}
    for wc in wanted_cols:
        match = next((c for c in df.columns if wc.lower() in c.lower()), None)
        # Specjalne traktowanie opisu, żeby nie pomylić z innymi
        if wc == "Katalog Członków CC - opis do 500 znaków":
             match = next((c for c in df.columns if "500" in c), None)
        # Specjalne traktowanie Photo/Logo
        if wc == "Photo": match = next((c for c in df.columns if "photo" in c.lower() and "nazwa" in c.lower()), None) or match
        if wc == "Logo": match = next((c for c in df.columns if "logo" in c.lower() and "nazwa" in c.lower()), None) or match
        
        if match: col_map[wc] = match

    # Przygotowanie DF do wyświetlenia
    display_cols = ["Imię", "Nazwisko", "Firma", "Branża", "Skala Biznesu", "Grupa CC"]
    # Upewniamy się, że mamy zmapowane kolumny
    final_display_cols = [col_map[c] for c in display_cols if c in col_map]
    
    # Kopia robocza
    df_view = df.copy()
    
    # Tworzenie kolumny sortującej (ukrytej) dla Skali Biznesu
    if "Skala Biznesu" in col_map:
        df_view["_sort_scale"] = df_view[col_map["Skala Biznesu"]].apply(parse_scale_business)

    # Domyślnie wszyscy zaznaczeni
    df_view.insert(0, "Wybierz", True)

    st.subheader("2. Lista do wygenerowania")
    st.info("Możesz sortować listę klikając w nagłówki kolumn (np. Skala Biznesu).")

    # EDYTOR DANYCH (TABELA)
    # Wyświetlamy tylko potrzebne kolumny + kolumnę do sortowania (którą ukryjemy wizualnie w configu)
    cols_to_show = ["Wybierz"] + final_display_cols
    if "_sort_scale" in df_view.columns:
        # Sortujemy wstępnie wg Skali (malejąco) jeśli istnieje, bo użytkownik o to pytał
        df_view = df_view.sort_values(by="_sort_scale", ascending=False)

    edited_df = st.data_editor(
        df_view,
        column_order=cols_to_show,
        hide_index=True,
        height=400,
        use_container_width=True
    )

    # 3. FILTROWANIE (POD LISTĄ)
    st.subheader("3. Filtrowanie Grup")
    
    if "Grupa CC" in col_map:
        group_col = col_map["Grupa CC"]
        all_groups = df[group_col].dropna().unique().tolist()
        # Domyślnie brak filtru (wszyscy), czy domyślnie wszyscy zaznaczeni?
        # User chciał wybierać dla jakiej grupy wygenerować.
        selected_groups = st.multiselect("Zaznacz grupy do uwzględnienia:", all_groups, default=all_groups)
        
        # Logika: bierzemy to co user "wyklikał" w tabeli (edited_df) I filtrujemy to grupami
        # Ważne: edited_df zawiera stan checkboxów "Wybierz".
        final_selection = edited_df[
            (edited_df["Wybierz"] == True) & 
            (edited_df[group_col].isin(selected_groups))
        ]
    else:
        final_selection = edited_df[edited_df["Wybierz"] == True]

    st.write(f"Wybrano **{len(final_selection)}** slajdów do wygenerowania.")

    # Ładowanie zdjęć
    images_map = {}
    if uploaded_zip:
        with zipfile.ZipFile(uploaded_zip) as z:
            for f in z.namelist():
                if not f.endswith('/') and "__MACOSX" not in f:
                    images_map[f.split('/')[-1].lower()] = z.read(f)

    # 4. GENEROWANIE
    st.markdown("---")
    if st.button("GENERUJ PREZENTACJĘ"):
        if final_selection.empty:
            st.error("Lista jest pusta. Zaznacz rekordy lub zmień filtry.")
        else:
            status_text = st.empty()
            progress_bar = st.progress(0)
            
            try:
                prs = Presentation(uploaded_pptx)
                # Zakładamy, że slajd nr 1 (indeks 0) to wzorzec
                
                # Iteracja po posortowanej przez użytkownika liście (final_selection)
                total = len(final_selection)
                for i, (idx, row) in enumerate(final_selection.iterrows()):
                    
                    # Klonowanie slajdu
                    new_slide = safe_duplicate_slide(prs, 0)
                    if new_slide is None:
                        continue # Pomijamy uszkodzony slajd
                    
                    # Przygotowanie danych (obsługa braku danych)
                    def get_val(key):
                        if key in col_map:
                            val = row.get(col_map[key])
                            return str(val) if pd.notna(val) else "-"
                        return "-"

                    replacements = {
                        "{Imię}": get_val("Imię"),
                        "{Nazwisko}": get_val("Nazwisko"),
                        "{Firma}": get_val("Firma"),
                        "{Branża}": get_val("Branża"),
                        "{Grupa CC}": get_val("Grupa CC"),
                        "{Skala Biznesu}": get_val("Skala Biznesu"),
                        "{Katalog Członków CC - opis do 500 znaków}": get_val("Katalog Członków CC - opis do 500 znaków")
                    }

                    # Podmiana na slajdzie
                    # Używamy list(), aby bezpiecznie modyfikować kolekcję w pętli
                    for shape in list(new_slide.shapes):
                        
                        # Teksty
                        replace_text_in_shape(shape, replacements)
                        
                        # Zdjęcia (po nazwie kształtu lub tekście w placeholderze)
                        shape_name_upper = shape.name.upper()
                        text_content = ""
                        if shape.has_text_frame:
                            text_content = shape.text_frame.text.strip().upper()

                        # PHOTO
                        if "PHOTO" in shape_name_upper or "PHOTO" in text_content:
                            photo_file = str(row.get(col_map.get("Photo"), "")).lower().strip()
                            if photo_file in images_map:
                                replace_image_in_shape(new_slide, shape, BytesIO(images_map[photo_file]))
                        
                        # LOGO
                        if "LOGO" in shape_name_upper or "LOGO" in text_content:
                            logo_file = str(row.get(col_map.get("Logo"), "")).lower().strip()
                            if logo_file in images_map:
                                replace_image_in_shape(new_slide, shape, BytesIO(images_map[logo_file]))

                    # Aktualizacja paska
                    progress_bar.progress((i + 1) / total)
                    status_text.text(f"Generuję: {replacements['{Nazwisko}']}")

                # Usuwamy slajd wzorcowy (pierwszy)
                xml_slides = prs.slides._sldIdLst
                slides_list = list(xml_slides)
                xml_slides.remove(slides_list[0])

                # Zapis
                output = BytesIO()
                prs.save(output)
                output.seek(0)
                
                timestamp = datetime.now().strftime("%H%M")
                file_name = f"Katalog_CC_{timestamp}.pptx"
                
                status_text.success("Gotowe! ✅")
                st.download_button(
                    label="POBIERZ PREZENTACJĘ",
                    data=output,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    type="primary"
                )

            except Exception as e:
                st.error(f"Wystąpił niespodziewany błąd: {e}")
                st.exception(e)

else:
    st.write("👈 Wgraj pliki w panelu bocznym.")
