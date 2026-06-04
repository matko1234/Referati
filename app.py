import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io

# Inicijalizacija memorije za literaturu i priloge
if "literatura" not in st.session_state:
    st.session_state.literatura = []
if "prilozi" not in st.session_state:
    st.session_state.prilozi = []

# Pomoćna funkcija koja dodaje nativni Wordov broj stranice (PAGE polje)
def dodaj_broj_stranice(run):
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)

def kreiraj_referat(naslov, predmet, mentor, ucenici, razred, datum, podnaslovi):
    doc = Document()

    # Postavljanje zadanog fonta za cijeli dokument na Times New Roman, 12pt
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    # --- AUTOMATSKO NUMERIRANJE (Sve stranice osim prve) ---
    section = doc.sections[0]
    section.different_first_page_header_footer = True  # Sakrij broj na naslovnici
    footer = section.footer
    p_footer = footer.paragraphs[0]
    p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_footer = p_footer.add_run()
    dodaj_broj_stranice(run_footer)

    # --- 1. NASLOVNICA (Precizno kalibrirani razmaci) ---
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_header = p_header.add_run("Gimnazija Karlovac")
    run_header.font.size = Pt(14)

    # Optimalan razmak do naslova
    for _ in range(4): doc.add_paragraph()

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_naslov = p_title.add_run(f"{naslov.upper()}\n")
    run_naslov.bold = True
    run_naslov.font.size = Pt(16)
    
    run_predmet = p_title.add_run(f"Referat iz {predmet}")
    run_predmet.font.size = Pt(14)

    # Optimalan razmak do podataka o učeniku/mentoru
    for _ in range(5): doc.add_paragraph()

    mentor_sufix = mentor if "prof." in mentor.lower() else f"{mentor}, prof."

    p_info = doc.add_paragraph()
    p_info.add_run(f"Mentor/ica: {mentor_sufix}\n")
    p_info.add_run(f"Učenik/ca: {ucenici}, {razred}")

    # Optimalan razmak do datuma na dnu prve stranice
    for _ in range(4): doc.add_paragraph()

    p_footer_text = doc.add_paragraph()
    p_footer_text.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_footer_text.add_run(f"Karlovac, {datum}.")
    
    doc.add_page_break()

    # --- 2. SADRŽAJ ---
    doc.add_heading('Sadržaj', level=1)
    popis_podnaslova = [p.strip() for p in podnaslovi.split('\n') if p.strip()]
    
    stranica = 3
    for p in popis_podnaslova:
        doc.add_paragraph(f"{p}\t{stranica}")
        stranica += 1
    
    if st.session_state.prilozi:
        doc.add_paragraph(f"Popis slikovnih priloga i tablica\t{stranica}")
        stranica += 1
        
    doc.add_paragraph(f"Popis literature i izvora\t{stranica}")
    doc.add_page_break()

    # --- 3. POGLAVLJA ---
    for p in popis_podnaslova:
        doc.add_heading(p, level=1)
        doc.add_page_break()

    # --- 4. ZASEBNA STRANICA ZA SLIKE I TABLICE (UPARENI I POMIČNI BLOKOVI) ---
    if st.session_state.prilozi:
        doc.add_heading("Popis slikovnih priloga i tablica", level=1)
        
        slika_cnt = 1
        tablica_cnt = 1
        
        for prilog in st.session_state.prilozi:
            # Kreiramo nevidljivu tablicu od 1 ćelije koja služi kao okvir/kontejner za grupiranje
            container_table = doc.add_table(rows=1, cols=1)
            
            # Pretvaramo tablicu u plutajući/slobodno pomični objekt oko kojeg tekst teče (Text Wrapping: Around)
            tblPr = container_table._tbl.tblPr
            tblpPr = OxmlElement('w:tblpPr')
            tblpPr.set(qn('w:leftFromText'), '180')
            tblpPr.set(qn('w:rightFromText'), '180')
            tblpPr.set(qn('w:topFromText'), '180')
            tblpPr.set(qn('w:bottomFromText'), '180')
            tblpPr.set(qn('w:vertAnchor'), 'paragraph')
            tblpPr.set(qn('w:horzAnchor'), 'margin')
            tblPr.insert(0, tblpPr)
            
            cell = container_table.cell(0, 0)
            
            # Unutar tog okvira dodajemo prateći tekst/analizu iz referata
            p_tekst = cell.paragraphs[0]
            if prilog['text']:
                p_tekst.add_run(f"Upareni tekst iz referata: {prilog['text']}")
            
            if prilog['type'] == "Slika/Graf":
                # Slika/Graf se umeće unutar okvira
                p_img = cell.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if prilog['image_bytes']:
                    p_img.add_run().add_picture(io.BytesIO(prilog['image_bytes']))
                
                # Potpis ide strogo ISPOD slike unutar istog okvira
                p_potpis = cell.add_paragraph()
                p_potpis.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_lbl = p_potpis.add_run(f"Slika {slika_cnt}. ")
                run_lbl.italic = True
                p_potpis.add_run(prilog['title'])
                if prilog['source']:
                    p_potpis.add_run(f" (Izvor: {prilog['source']})")
                slika_cnt += 1
                
            elif prilog['type'] == "Tablica":
                # Za tablicu, potpis i naslov idu IZNAD prikaza unutar okvira
                p_potpis = cell.add_paragraph()
                p_potpis.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_lbl = p_potpis.add_run(f"Tablica {tablica_cnt}")
                run_lbl.italic = True
                
                p_naslov = cell.add_paragraph()
                p_naslov.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_naslov.add_run(prilog['title'])
                
                # Tablica (slika) unutar okvira
                p_img = cell.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if prilog['image_bytes']:
                    p_img.add_run().add_picture(io.BytesIO(prilog['image_bytes']))
                
                if prilog['source']:
                    p_izvor = cell.add_paragraph()
                    p_izvor.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_izvor.add_run(f"Izvor: {prilog['source']}")
                tablica_cnt += 1
            
            # Razmak izvan okvira, da se sljedeći prilozi ne lijepe
            doc.add_paragraph("\n")
            
        doc.add_page_break()

    # --- 5. LITERATURA ---
    doc.add_heading("Popis literature i izvora", level=1)
    
    if st.session_state.literatura:
        for izvor in sorted(st.session_state.literatura):
            doc.add_paragraph(izvor)
    else:
        doc.add_paragraph("(Učenik nije unio literaturu u aplikaciji.)")

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- WEB SUČELJE ---
st.title("Generator školskog referata")

st.header("1. Osnovni podaci")
naslov = st.text_input("Naslov referata")
predmet = st.text_input("Naziv predmeta")
mentor = st.text_input("Ime i prezime mentora")
ucenici = st.text_input("Ime i prezime učenika")
razred = st.text_input("Razred")
datum = st.text_input("Datum")

st.header("2. Izborne opcije")

st.subheader("Sadržaj")
podnaslovi_input = st.text_area(
    "Unesi podnaslove referata (svaki u novi red):", 
    "Uvod\nRazrada teme\nZaključak"
)

# --- SEKCIJA ZA UPIRIVANJE GRAFOVA I TABLICA ---
st.subheader("Slikovni prilozi i Tablice")
st.write("Ovdje dodajte grafove, slike ili tablice uparene s tekstom. U Wordu će se stvoriti slobodno pomični blokovi.")

tip_priloga = st.selectbox("Odaberi tip priloga:", ["Slika/Graf", "Tablica"])
naziv_priloga = st.text_input("Naziv/Naslov priloga (npr. Grafički prikaz... ili Broj učenika...)")
tekst_priloga = st.text_area("Unesi tekst iz referata koji se odnosi na ovaj prilog:")
izvor_priloga = st.text_input("Izvor priloga (ostavi prazno ako je tvoj vlastiti rad)")
ucitana_slika = st.file_uploader("Učitaj slikovni prikaz priloga (PNG, JPG, JPEG):", type=["png", "jpg", "jpeg"])

if st.button("➕ Dodaj prilog u dokument"):
    if naziv_priloga:
        img_bytes = ucitana_slika.read() if ucitana_slika is not None else None
        st.session_state.prilozi.append({
            "type": tip_priloga,
            "title": naziv_priloga,
            "text": tekst_priloga,
            "source": izvor_priloga,
            "image_bytes": img_bytes
        })
        st.success(f"{tip_priloga} je uspješno spremljen, uparen i postavljen kao slobodno pomičan!")
    else:
        st.error("Molimo unesite barem naziv priloga.")

if st.session_state.prilozi:
    st.write(f"**Trenutno dodano priloga: {len(st.session_state.prilozi)}**")
    if st.button("🗑️ Obriši sve priloge"):
        st.session_state.prilozi = []
        st.rerun()

# --- SEKCIJA ZA LITERATURU ---
st.subheader("Literatura i izvori")
tip_izvora = st.selectbox("Odaberi vrstu izvora za unos:", ["Knjiga", "Web stranica"])

col1, col2 = st.columns(2)
with col1:
    autor = st.text_input("Autor (Prezime, Ime)")
    godina = st.text_input("Godina izdanja")
    naslov_izvora = st.text_input("Naslov djela/članka")

with col2:
    if tip_izvora == "Knjiga":
        izdavac = st.text_input("Izdavač")
        grad = st.text_input("Mjesto izdanja")
    else:
        url = st.text_input("URL (poveznica)")
        datum_pristupa = st.text_input("Datum pristupa")

if st.button("➕ Dodaj izvor u literaturu"):
    if tip_izvora == "Knjiga" and autor and naslov_izvora:
        izvor_str = f"{autor}. {godina + '. ' if godina else ''}{naslov_izvora}. {izdavac}. {grad}."
        st.session_state.literatura.append(izvor_str)
        st.success("Knjiga je dodana!")
    elif tip_izvora == "Web stranica" and naslov_izvora and url:
        godina_str = f"{godina}. " if font else ""
        izvor_str = f"{autor}. {godina_str}{naslov_izvora}. Preuzeto s {url} (pristupljeno {datum_pristupa})."
        st.session_state.literatura.append(izvor_str)
        st.success("Web stranica je dodana!")
    else:
        st.error("Molimo ispuni barem Autora i Naslov izvora.")

if st.session_state.literatura:
    st.write("**Trenutno dodani izvori:**")
    for lit in sorted(st.session_state.literatura):
        st.write(f"- {lit}")
    if st.button("🗑️ Obriši svu literaturu"):
        st.session_state.literatura = []
        st.rerun()

st.divider()

if st.button("Pripremi dokument"):
    if naslov and predmet and ucenici:
        word_file = kreiraj_referat(naslov, predmet, mentor, ucenici, razred, datum, podnaslovi_input)
        ime_datoteke = f"Referat_{ucenici.replace(' ', '_')}.docx"
        
        st.success("Dokument je spreman!")
        st.download_button(
            label="⬇️ Preuzmi Referat",
            data=word_file,
            file_name=ime_datoteke,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.error("Ispuni barem Naslov, Predmet i Ime učenika!")
