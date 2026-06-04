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

    # --- 1. NASLOVNICA ---
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_header = p_header.add_run("Gimnazija Karlovac")
    run_header.font.size = Pt(14)

    for _ in range(7): doc.add_paragraph()

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_naslov = p_title.add_run(f"{naslov.upper()}\n")
    run_naslov.bold = True
    run_naslov.font.size = Pt(16)
    
    run_predmet = p_title.add_run(f"Referat iz {predmet}")
    run_predmet.font.size = Pt(14)

    for _ in range(8): doc.add_paragraph()

    mentor_sufix = mentor if "prof." in mentor.lower() else f"{mentor}, prof."

    p_info = doc.add_paragraph()
    p_info.add_run(f"Mentor/ica: {mentor_sufix}\n")
    p_info.add_run(f"Učenik/ca: {ucenici}, {razred}")

    for _ in range(6): doc.add_paragraph()

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

    # --- 4. ZASEBNA STRANICA ZA SLIKE I TABLICE ---
    if st.session_state.prilozi:
        doc.add_heading("Popis slikovnih priloga i tablica", level=1)
        
        slika_cnt = 1
        tablica_cnt = 1
        
        for prilog in st.session_state.prilozi:
            # Prvo ispisujemo upareni tekst/analizu koju je učenik unio
            if prilog['text']:
                p_tekst = doc.add_paragraph()
                p_tekst.add_run(f"Upareni tekst iz referata: {prilog['text']}")
            
            if prilog['type'] == "Slika/Graf":
                # Slika ide PRVA
                if prilog['image_bytes']:
                    doc.add_picture(io.BytesIO(prilog['image_bytes']))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Potpis ide ISPOD slike [cite: 46]
                p_potpis = doc.add_paragraph()
                p_potpis.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_lbl = p_potpis.add_run(f"Slika {slika_cnt}. ")  # Točka iza broja [cite: 48]
                run_lbl.italic = True  # Ukošeno [cite: 47]
                p_potpis.add_run(prilog['title'])
                if prilog['source']:
                    p_potpis.add_run(f" (Izvor: {prilog['source']})")  # Izvor ako postoji [cite: 49]
                slika_cnt += 1
                
            elif prilog['type'] == "Tablica":
                # Potpis i naslov idu IZNAD tablice [cite: 56]
                p_potpis = doc.add_paragraph()
                p_potpis.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_lbl = p_potpis.add_run(f"Tablica {tablica_cnt}")  # Bez točke iza broja [cite: 58]
                run_lbl.italic = True  # Ukošeno [cite: 57]
                
                p_naslov = doc.add_paragraph()
                p_naslov.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_naslov.add_run(prilog['title'])  # Naslov u zasebnom redu [cite: 60]
                
                # Tablica (u obliku slike) [cite: 61]
                if prilog['image_bytes']:
                    doc.add_picture(io.BytesIO(prilog['image_bytes']))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                if prilog['source']:
                    p_izvor = doc.add_paragraph()
                    p_izvor.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_izvor.add_run(f"Izvor: {prilog['source']}")  # Izvor ispod tablice [cite: 63]
                tablica_cnt += 1
                
            # Razmak između različitih priloga
            doc.add_paragraph("\n--------------------------------------------------\n")
            
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
predmet = st.text_input("Naziv predmeta (npr. povijesti)")
mentor = st.text_input("Ime i prezime mentora (titula prof. se dodaje automatski)")
ucenici = st.text_input("Ime i prezime učenika")
razred = st.text_input("Razred (npr. 3.a)")
datum = st.text_input("Datum (npr. 3. rujna 2018.)")

st.header("2. Izborne opcije")

st.subheader("Sadržaj")
podnaslovi_input = st.text_area(
    "Unesi podnaslove referata (svaki u novi red):", 
    "Uvod\nRazrada teme\nZaključak"
)

# --- SEKCIJA ZA UPIRIVANJE GRAFOVA I TABLICA ---
st.subheader("Slikovni prilozi i Tablice")
st.write("Ovdje možete dodati grafove, slike ili tablice uparene s tekstom koje će se pravilno formatirati na kraju rada.")

tip_priloga = st.selectbox("Odaberi tip priloga:", ["Slika/Graf", "Tablica"])
naziv_priloga = st.text_input("Naziv/Naslov priloga (npr. Grafički prikaz... ili Broj učenika...)")
tekst_priloga = st.text_area("Unesi tekst iz referata koji se odnosi na ovaj prilog (objašnjenje/analiza):")
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
        st.success(f"{tip_priloga} je uspješno spremljen i uparen!")
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
    autor = st.text_input("Autor (Prezime, Ime ili dr.)")
    godina = st.text_input("Godina izdanja")
    naslov_izvora = st.text_input("Naslov djela/članka")

with col2:
    if tip_izvora == "Knjiga":
        izdavac = st.text_input("Izdavač (npr. Školska knjiga)")
        grad = st.text_input("Mjesto izdanja (npr. Zagreb)")
    else:
        url = st.text_input("URL (poveznica)")
        datum_pristupa = st.text_input("Datum pristupa (npr. 31. kolovoza 2018.)")

if st.button("➕ Dodaj izvor u literaturu"):
    if tip_izvora == "Knjiga" and autor and naslov_izvora:
        izvor_str = f"{autor}. {godina + '. ' if godina else ''}{naslov_izvora}. {izdavac}. {grad}."
        st.session_state.literatura.append(izvor_str)
        st.success("Knjiga je dodana!")
    elif tip_izvora == "Web stranica" and naslov_izvora and url:
        godina_str = f"{godina}. " if godina else ""
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
