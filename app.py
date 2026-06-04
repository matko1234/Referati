import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

if "literatura" not in st.session_state:
    st.session_state.literatura = []

def kreiraj_referat(naslov, predmet, mentor, ucenici, razred, datum, podnaslovi):
    doc = Document()

    # Postavljanje zadanog fonta za cijeli dokument na Times New Roman, 12pt
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    # --- 1. NASLOVNICA ---
    # Gimnazija Karlovac
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_header = p_header.add_run("Gimnazija Karlovac")
    run_header.font.size = Pt(14)

    # Razmak do naslova
    for _ in range(7): doc.add_paragraph()

    # Naslov teme i predmet
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_naslov = p_title.add_run(f"{naslov.upper()}\n")
    run_naslov.bold = True
    run_naslov.font.size = Pt(16)
    
    run_predmet = p_title.add_run(f"Referat iz {predmet}")
    run_predmet.font.size = Pt(14)

    # Razmak do podataka o mentoru i učeniku
    for _ in range(8): doc.add_paragraph()

    # Automatsko dodavanje "prof." ako već nije upisano
    mentor_sufix = mentor if "prof." in mentor.lower() else f"{mentor}, prof."

    p_info = doc.add_paragraph()
    p_info.add_run(f"Mentor/ica: {mentor_sufix}\n")
    p_info.add_run(f"Učenik/ca: {ucenici}, {razred}")

    # Razmak do dna stranice
    for _ in range(6): doc.add_paragraph()

    p_footer = doc.add_paragraph()
    p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_footer.add_run(f"Karlovac, {datum}.")
    
    doc.add_page_break()

    # --- 2. SADRŽAJ ---
    doc.add_heading('Sadržaj', level=1)
    popis_podnaslova = [p.strip() for p in podnaslovi.split('\n') if p.strip()]
    
    stranica = 3
    for p in popis_podnaslova:
        doc.add_paragraph(f"{p}\t{stranica}")
        stranica += 1
    doc.add_paragraph(f"Popis literature i izvora\t{stranica}")
    doc.add_page_break()

    # --- 3. POGLAVLJA ---
    for p in popis_podnaslova:
        doc.add_heading(p, level=1)
        doc.add_page_break()

    # --- 4. LITERATURA ---
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
predmet = st.text_input("Referat iz...")
mentor = st.text_input("Ime i prezime mentora (titula prof. dodaje se automatski)")
ucenici = st.text_input("Ime i prezime učenika")
razred = st.text_input("Razred)")
datum = st.text_input("Datum")

st.header("2. Izborne opcije")

st.subheader("Sadržaj")
podnaslovi_input = st.text_area(
    "Unesi podnaslove referata (svaki u novi red):", 
    "Uvod\nRazrada teme\nZaključak"
)

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
        godina_str = f"{godina}. " if godina else ""
        izvor_str = f"{autor}. {godina_str}{naslov_izvora}. Preuzeto s {url} (pristupljeno {datum_pristupa})."
        st.session_state.literatura.append(izvor_str)
        st.success("Web stranica je dodana!")
    else:
        st.error("Molimo ispuni barem Autora i Naslov.")

if st.session_state.literatura:
    st.write("**Trenutno dodani izvori (bit će automatski sortirani abecedno):**")
    for lit in st.session_state.literatura:
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
