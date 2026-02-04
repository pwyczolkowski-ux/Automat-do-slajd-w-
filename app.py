import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io
import zipfile
from PIL import Image

# --- KONFIGURACJA STRONY ---
st.set_page_config(page_title="Generator Katalogu CC", layout="wide")

# Wymuszamy style CSS dla pewności (czarny tekst na białym tle)
st.markdown("""
    <style>
    .stApp {
        background-color: white;
        color: black;
    }
    div[data-testid="stDataFrame"] {
        background-color: white;
        border: 1px solid #ddd;
    }
    h1, h2, h3, p, label {
        color: black !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- FUNKCJE POMOCNICZE ---

def crop_image_to_aspect_ratio(image_bytes, target_ratio):
    """
    Przycina zdjęcie w pamięci do zadanego formatu (np. kwadrat, prostokąt),
    aby uniknąć efektu rozciągnięcia (działa jak object-fit: cover).
    """
    with Image.open(io.BytesIO(image_bytes)) as img:
        img_ratio = img.width / img.height
        
        if img_ratio > target_ratio:
            # Zdjęcie jest za szerokie - ucinamy boki
            new_width = int(target_ratio * img.height)
            left = (img.width - new_width) / 2
            top = 0
            right = (img.width + new_width) / 2
            bottom = img.height
        else:
            # Zdjęcie jest za wysokie - ucinamy górę/dół
            new_height = int(img.width / target_ratio)
            left = 0
            top = (img.height - new_height) / 2
            right = img.width
            bottom = (img.height + new_height) / 2
            
        img = img.crop((left, top, right, bottom))
        
        output = io.BytesIO()
        img.save(output, format=img.format if img.format else 'JPEG')
        return output

def find_image_in_zip(zip_file, filename_base):
    """Szuka pliku w ZIP ignorując wielkość liter i rozszerzenie."""
    # filename_base to np. "Jan_Kowalski"
    for name in zip_file.namelist():
        # Ignorujemy foldery (MacOS tworzy __MACOSX)
        if name.startswith("__MACOSX") or name.endswith("/"):
            continue
            
        # Sprawdzamy czy nazwa pliku (bez rozszerzenia) pasuje
        clean_name = name.split('/')[-1] # usuwa ścieżkę folderów
        name_no_ext = clean_name.rsplit('.', 1)[0]
        
        if name_no_ext.lower() == filename_base.lower():
            return zip_file.read(name)
    return None

def generate_pptx(df, pptx_template, images_zip):
    prs = Presentation(pptx_template)
    
    # Zakładamy, że używamy pierwszego układu slajdu we wzorcu (indeks 0 lub 1)
    # Warto sprawdzić w PPTX, który to layout. Tu przyjmuję Layout nr 1 (często "Title and Content")
    # Jeśli Twój szablon jest niestandardowy, zmień index np. na prs.slide_layouts[0]
    slide_layout = prs.slide_layouts[0] 

    # Otwieramy ZIP raz
    z = zipfile.ZipFile(images_zip)

    # Pasek postępu
    progress_bar = st.progress(0)
    total_rows = len(df)

    for index, row in df.iterrows():
        # Tworzymy nowy slajd
        slide = prs.slides.add_slide(slide_layout)
        
        # Przygotowanie danych (obsługa braków danych - fillna)
        imie = str(row.get('Imię', ''))
        nazwisko = str(row.get('Nazwisko', ''))
        firma = str(row.get('Firma', ''))
        opis = str(row.get('Opis', 'Brak opisu'))
        skala = str(row.get('Skala', ''))
        
        # Nazwa pliku zdjęcia jakiej szukamy (np. Jan_Kowalski)
        foto_szukane = f"{imie}_{nazwisko}".strip().replace(" ", "_")

        # Iterujemy po kształtach (placeholderach) na slajdzie
        for shape in slide.placeholders:
            # Używamy nazw nadanych w Selection Pane (Okienko zaznaczenia)
            
            if shape.name == "DANE_OSOBOWE":
                shape.text = f"{imie} {nazwisko}"
            
            elif shape.name == "FIRMA_BOX":
                shape.text = firma
                
            elif shape.name == "SKALA_BOX":
                shape.text = skala
                
            elif shape.name == "OPIS_BOX":
                shape.text = opis
                # Opcjonalnie formatowanie tekstu opisu
                if shape.has_text_frame:
                    p = shape.text_frame.paragraphs[0]
                    p.font.size = Pt(11)
                    p.font.color.rgb = RGBColor(0, 0, 0)

            elif shape.name == "FOTO_BOX":
                # Szukamy zdjęcia
                img_data = find_image_in_zip(z, foto_szukane)
                if img_data:
                    # Obliczamy proporcje placeholdera
                    target_ratio = shape.width / shape.height
                    # Przycinamy zdjęcie (crop)
                    cropped_img = crop_image_to_aspect_ratio(img_data, target_ratio)
                    
                    # Wstawiamy zdjęcie w placeholder
                    # insert_picture automatycznie zastępuje placeholder zachowując jego pozycję
                    shape.insert_picture(cropped_img)
                else:
                    # Jeśli brak zdjęcia, można wpisać tekst
                    shape.text = "Brak zdjęcia"

        progress_bar.progress((index + 1) / total_rows)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# --- INTERFEJS UŻYTKOWNIKA ---

st.title("Generowanie Katalogu CC (High Contrast)")

st.subheader("1. Wgraj pliki")

col1, col2, col3 = st.columns(3)

with col1:
    uploaded_excel = st.file_uploader("Baza Danych (.xlsx)", type=['xlsx'])
with col2:
    uploaded_pptx = st.file_uploader("Szablon (.pptx)", type=['pptx'])
with col3:
    uploaded_zip = st.file_uploader("Zdjęcia (.zip)", type=['zip'])

if uploaded_excel and uploaded_pptx and uploaded_zip:
    try:
        # Wczytanie Excela
        df = pd.read_excel(uploaded_excel)

        # -- NAPRAWA NAZW KOLUMN --
        # Mapujemy twoją długą nazwę na prostą "Opis"
        rename_map = {
            "Katalog Członków CC - opis do 500 znaków": "Opis",
            "Skala Biznesu": "Skala",
            # Upewnij się, że te kolumny istnieją w Excelu (Imię, Nazwisko, Firma)
        }
        df = df.rename(columns=rename_map)
        
        # Filtrujemy tylko te kolumny, które nas obchodzą
        # Używamy .get, żeby kod się nie wywalił jak czegoś brakuje
        required_cols = ['Imię', 'Nazwisko', 'Firma', 'Opis', 'Skala']
        
        # Sprawdzenie czy kolumny istnieją
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            st.error(f"Brakuje w Excelu kolumn: {missing}. Sprawdź nazwy!")
        else:
            st.success("Pliki wczytane poprawnie.")
            
            # Podgląd danych (wymuszony jasny motyw w CSS wyżej)
            st.subheader("2. Podgląd danych (Pierwsze 5 wierszy)")
            st.dataframe(df[required_cols].head())
            
            st.subheader("3. Generowanie")
            if st.button("Generuj Katalog PowerPoint"):
                with st.spinner("Przetwarzanie slajdów i przycinanie zdjęć..."):
                    try:
                        out_file = generate_pptx(df, uploaded_pptx, uploaded_zip)
                        
                        st.download_button(
                            label="📥 Pobierz gotowy plik .pptx",
                            data=out_file,
                            file_name="Katalog_CC_Generated.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                        st.success("Gotowe!")
                    except Exception as e:
                        st.error(f"Wystąpił błąd podczas generowania: {e}")
                        st.info("Wskazówka: Sprawdź czy we Wzorcu Slajdów nazwy placeholderów to dokładnie: FOTO_BOX, OPIS_BOX, DANE_OSOBOWE, etc.")

    except Exception as e:
        st.error(f"Błąd pliku Excel: {e}")

else:
    st.info("Wgraj wszystkie 3 pliki, aby rozpocząć.")
