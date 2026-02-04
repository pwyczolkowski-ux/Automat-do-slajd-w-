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

# Wymuszenie stylu "Clean White" (Białe tło, czarne litery, brak kolorowych ozdobników)
st.markdown("""
    <style>
        /* Reset kolorów systemowych Streamlit */
        .stApp {
            background-color: #FFFFFF;
            color: #000000;
        }
        /* Nagłówki */
        h1, h2, h3, h4, h5, h6, p, label, .stMarkdown {
            color: #000000 !important;
            font-family: 'Roboto', sans-serif !important;
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
        }
        /* Inputy i tabele */
        .stDataFrame, .stDataEditor {
            border: 1px solid #000000;
        }
        [data-testid="stSidebar"] {
            background-color: #F0F0F0;
            border-right: 1px solid #000000;
        }
        hr {
            border-color: #000000;
        }
        /* Ukrycie domyślnego menu Streamlit */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# --- FUNKCJE LOGICZNE ---

def duplicate_slide(pres, index):
    """
    Klonuje slajd o podanym indeksie w tej samej prezentacji.
    To kluczowa funkcja naprawiająca problem "pustych slajdów".
    """
    source = pres.slides[index]
    blank_slide_layout = pres.slide_layouts[6] # Pusty layout
    dest = pres.slides.add_slide(blank_slide_layout)

    # Kopiowanie elementów (kształtów) ze źródła do celu
    for shape in source.shapes:
        new_el = copy.deepcopy(shape.element)
        dest.shapes._spTree.insert_element_before(new_el, 'p:extLst')

    # Odświeżenie relacji (aby style zadziałały)
    for key, value in source.part.rels.items():
        if "notesSlide" not in value.reltype:
            dest.part.rels.add_relationship(
                value.reltype,
                value._target,
                value.rId
            )
    return dest

def clean_polish_typography(text):
    """Zapobiega wiszącym spójnikom."""
    if not isinstance(text, str): return text
    conjunctions = [" w ", " z ", " i ", " a ", " o ", " u ", " na ", " do "]
    for word in conjunctions:
        pattern = re.compile(re.escape(word), re.IGNORECASE)
        text = pattern.sub(lambda m: m.group(0).replace(' ', '\u00A0', 1), text)
    return text

def replace_text_in_shape(shape, replacements):
    """Szuka i zamienia tekst wewnątrz kształtu (TextFrame)."""
    if not shape.has_text_frame:
        return

    # Iterujemy po akapitach i fragmentach tekstu
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            for key, value in replacements.items():
                if key in run.text:
                    # Podmiana tekstu
                    new_text = run.text.replace(key, str(value))
                    run.text = new_text
                    
                    # Logika zmniejszania czcionki dla długich opisów (tylko dla Opisu)
                    if key == "{Katalog Członków CC - opis do 500 znaków}":
                        clean_text = clean_polish_typography(str(value))
                        run.text = run.text.replace(str(value), clean_text) # Aplikujemy typografię
                        if len(str(value)) > 600:
                            run.font.size = Pt(8)
                        elif len(str(value)) > 450:
                            run.font.size = Pt(9)

def replace_image_in_shape(slide, shape, image_stream):
    """Podmienia kształt o nazwie/tekście PHOTO lub LOGO na obrazek."""
    try:
        # Pobieramy wymiary starego kształtu
        left, top = shape.left, shape.top
        width, height = shape.width, shape.height
        
        # Wstawiamy nowy obrazek w to samo miejsce
        slide.shapes.add_picture(image_stream, left, top, width, height)
        
        # Usuwamy stary kształt (placeholder)
        # Hack: przesuwamy stary kształt poza slajd lub usuwamy z XML
        sp = shape._element
        sp.getparent().remove(sp)
    except Exception as e:
        print(f"Błąd obrazka: {e}")

# --- INTERFEJS ---

st.title("Generator Katalogu CC")
st.markdown("---")

# 1. SIDEBAR - PLIKI
with st.sidebar:
    st.header("Wgraj pliki")
    uploaded_excel = st.file_uploader("1. Excel (Dane)", type=['xlsx', 'csv'])
    uploaded_pptx = st.file_uploader("2. Szablon (.pptx)", type=['pptx'])
    uploaded_zip = st.file_uploader("3. Zdjęcia (.zip)", type=['zip'])
    
    # Przycisk odwrócenia kontrastu (prostym hackiem CSS - opcjonalnie)
    st.markdown("---")
    st.markdown("**Ustawienia:**")
    st.caption("Sortowanie i filtrowanie dostępne po wgraniu plików.")

if uploaded_excel and uploaded_pptx:
    # 2. PRZETWARZANIE DANYCH
    try:
        if uploaded_excel.name.endswith('.csv'):
            df = pd.read_csv(uploaded_excel)
        else:
            df = pd.read_excel(uploaded_excel)
        
        # Usuwamy spacje z nazw kolumn
        df.columns = df.columns.str.strip()
        
        # WYMAGANE KOLUMNY (Tylko te importujemy do widoku)
        wanted_columns = [
            "Imię", 
            "Nazwisko", 
            "Firma", 
            "Branża", 
            "Katalog Członków CC - opis do 500 znaków", 
            "Grupa CC",
            "Photo", # Potrzebne do logiki, ale nie musimy wyświetlać w edytorze jeśli nie chcesz
            "Logo"
        ]
        
        # Mapowanie nazw (jeśli w pliku są "Photo nazwa pliku" zamiast "Photo")
        # Prosta logika szukania odpowiedników
        final_cols = []
        for w_col in wanted_columns:
            found = False
            for df_col in df.columns:
                if w_col.lower() in df_col.lower() and "opis" not in df_col.lower(): # Unikamy pomyłki przy Photo/Opis
                    final_cols.append(df_col)
                    found = True
                    break
                # Specjalny przypadek dla długiego opisu
                if "opis" in w_col.lower() and "opis" in df_col.lower() and "500" in df_col.lower():
                    final_cols.append(df_col)
                    found = True
                    break
            if not found:
                # Jeśli nie znaleziono idealnego dopasowania, szukamy luźniej lub zostawiamy
                pass

        # Filtrujemy DF do wymaganych kolumn (plus te, które udało się znaleźć)
        # Dla bezpieczeństwa bierzemy te, które na pewno są
        valid_cols = [c for c in wanted_columns if c in df.columns]
        
        # Jeśli brakuje kluczowych, próbujemy mapować ręcznie dla Twoich plików
        # (Hardcode pod Twoje pliki CSV, żeby zawsze działało)
        if "Photo nazwa pliku" in df.columns: 
            df["Photo"] = df["Photo nazwa pliku"]
            valid_cols.append("Photo")
        if "Logo nazwa pliku" in df.columns: 
            df["Logo"] = df["Logo nazwa pliku"]
            valid_cols.append("Logo")

        # Tworzymy czysty widok
        display_cols = ["Imię", "Nazwisko", "Firma", "Branża", "Grupa CC"]
        # Sprawdzamy czy istnieją w df
        display_cols = [c for c in display_cols if c in df.columns]
        
        # Sortowanie i Filtrowanie
        col_L, col_R = st.columns([1, 2])
        
        with col_L:
            st.subheader("Filtrowanie")
            if "Grupa CC" in df.columns:
                all_groups = df["Grupa CC"].dropna().unique().tolist()
                selected_groups = st.multiselect("Wybierz Grupy:", all_groups, default=all_groups)
                df_filtered = df[df["Grupa CC"].isin(selected_groups)].copy()
            else:
                df_filtered = df.copy()

        with col_R:
            st.subheader("Lista do wygenerowania")
            # Sortowanie
            sort_mode = st.selectbox("Sortuj według:", ["Domyślne", "Nazwisko A-Z", "Firma A-Z"])
            if sort_mode == "Nazwisko A-Z" and "Nazwisko" in df_filtered.columns:
                df_filtered = df_filtered.sort_values("Nazwisko")
            elif sort_mode == "Firma A-Z" and "Firma" in df_filtered.columns:
                df_filtered = df_filtered.sort_values("Firma")

            # Dodajemy kolumnę "Wybierz"
            df_filtered.insert(0, "Wybierz", True)
            
            # Edytor danych (Tylko wybrane kolumny widoczne)
            edited_df = st.data_editor(
                df_filtered[["Wybierz"] + display_cols], # Pokazujemy tylko proste kolumny
                hide_index=True,
                height=300,
                use_container_width=True
            )
            
            # Pobieramy ID wybranych wierszy (indeksy z oryginalnego DF_filtered)
            # Ponieważ data_editor zwraca zmodyfikowany DF, musimy połączyć go z resztą danych (Photo, Logo, Opis)
            # Najbezpieczniej: bierzemy indeksy z edited_df gdzie Wybierz=True i filtrujemy df_filtered
            selected_indices = edited_df[edited_df["Wybierz"] == True].index
            final_data = df_filtered.loc[selected_indices]

        # Ładowanie ZIP ze zdjęciami
        images_map = {}
        if uploaded_zip:
            with zipfile.ZipFile(uploaded_zip) as z:
                for f in z.namelist():
                    if not f.endswith('/'): 
                         # Klucz: sama nazwa pliku małymi literami
                        images_map[f.split('/')[-1].lower()] = z.read(f)

        # 3. GENEROWANIE
        st.markdown("---")
        if st.button("GENERUJ SLAJDY"):
            if final_data.empty:
                st.error("Nie wybrano żadnych osób.")
            else:
                prs = Presentation(uploaded_pptx)
                
                # Zamiast używać "Layoutu", KLONUJEMY pierwszy slajd
                # Zakładamy, że slajd 0 to Twój wzorzec narysowany ręcznie
                template_slide_index = 0 
                
                progress_bar = st.progress(0)
                
                for i, (idx, row) in enumerate(final_data.iterrows()):
                    # 1. Sklonuj slajd wzorcowy
                    new_slide = duplicate_slide(prs, template_slide_index)
                    
                    # 2. Przygotuj dane do podmiany
                    # Pobieramy bezpiecznie, nawet jak kolumna nie istnieje
                    def val(col_name):
                        if col_name in df.columns:
                            v = row[col_name]
                            return str(v) if pd.notna(v) else "-"
                        return "-"

                    replacements = {
                        "{Imię}": val("Imię"),
                        "{Nazwisko}": val("Nazwisko"),
                        "{Firma}": val("Firma"),
                        "{Branża}": val("Branża"),
                        "{Grupa CC}": val("Grupa CC"),
                        "{Skala Biznesu}": val("Skala Biznesu"), # Jeśli jest w excelu
                        "{Katalog Członków CC - opis do 500 znaków}": val("Katalog Członków CC - opis do 500 znaków")
                    }

                    # 3. Iterujemy po kształtach na NOWYM slajdzie i podmieniamy
                    # Musimy użyć list(new_slide.shapes), bo będziemy usuwać niektóre (przy podmianie zdjęć)
                    for shape in list(new_slide.shapes):
                        
                        # A. Podmiana Tekstu
                        if shape.has_text_frame:
                            replace_text_in_shape(shape, replacements)
                            
                            # Sprawdzamy czy to placeholder tekstowy PHOTO/LOGO (jeśli user wpisał tekst zamiast Alt Text)
                            txt = shape.text_frame.text.strip()
                            if txt == "PHOTO" or txt == "{PHOTO}":
                                photo_name = str(row.get("Photo", "")).lower().strip()
                                if photo_name in images_map:
                                    replace_image_in_shape(new_slide, shape, BytesIO(images_map[photo_name]))
                                    
                            elif txt == "LOGO" or txt == "{LOGO}":
                                logo_name = str(row.get("Logo", "")).lower().strip()
                                if logo_name in images_map:
                                    replace_image_in_shape(new_slide, shape, BytesIO(images_map[logo_name]))

                        # B. Podmiana po nazwie kształtu (Selection Pane)
                        # Jeśli nazwałeś kształt np. "PHOTO_PLACEHOLDER"
                        if "PHOTO" in shape.name.upper():
                            photo_name = str(row.get("Photo", "")).lower().strip()
                            if photo_name in images_map:
                                replace_image_in_shape(new_slide, shape, BytesIO(images_map[photo_name]))
                        
                        if "LOGO" in shape.name.upper():
                            logo_name = str(row.get("Logo", "")).lower().strip()
                            if logo_name in images_map:
                                replace_image_in_shape(new_slide, shape, BytesIO(images_map[logo_name]))

                    progress_bar.progress((i + 1) / len(final_data))

                # Na koniec usuwamy slajd wzorcowy (pierwszy), żeby nie było go w wynikowym pliku
                # (Hack na usunięcie slajdu w python-pptx)
                xml_slides = prs.slides._sldIdLst
                slides_list = list(xml_slides)
                xml_slides.remove(slides_list[0])

                # Zapis
                output = BytesIO()
                prs.save(output)
                output.seek(0)
                
                timestamp = datetime.now().strftime("%Y%m%d_%H%M")
                st.success("Gotowe!")
                st.download_button(
                    "POBIERZ PLIK .PPTX",
                    data=output,
                    file_name=f"Katalog_CC_{timestamp}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

    except Exception as e:
        st.error(f"Wystąpił błąd: {e}")
        st.write("Sprawdź czy nazwy kolumn w Excelu są poprawne.")

else:
    st.info("👈 Wgraj pliki w menu po lewej stronie.")
