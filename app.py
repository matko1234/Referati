import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

def kreiraj_referat(naslov, predmet, mentor, ucenici, razred, datum):
    doc = Document()

    # Naslovnica
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_header.add_run("Gimnazija Karlovac\n\n\n\n\n\n\n")

    run_naslov = p_header.add_run(f"{naslov.upper()}\n")
    run_naslov.bold = True
    run_naslov.font.size = Pt(16)
    p_header.add_run(f"Referat iz {predmet}\n\n\n\n\n\n\n\n\n")

    p_info = doc.add_paragraph()
    p_info.add_run(f"Mentor/ica: {mentor}\n")
    p_info.add_run(f"Učenik/ca: {ucenici}, {razred}\n\n\n\n\n\n")

    p_footer = doc.add_paragraph()
    p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_footer.add_run(f"Karlovac, {datum}.")
    
    doc.add_page_break()

    # Sadržaj
    doc.add_heading('Sadržaj', level=1)
    doc.add_paragraph("Uvod\t3\nRazrada teme\t4\nZaključak\t7\nPopis literature i izvora\t8")
    doc.add_page_break()

    # Poglavlja
    poglavlja = ["Uvod", "Razrada teme", "Zaključak", "Popis literature i izvora"]
    for poglavlje in poglavlja:
        doc.add_heading(poglavlje, level=1)
        if poglavlje == "Popis literature i izvora":
            doc.add_paragraph(
                "U popisu literature bibliografske jedinice navode se abecednim redom.\n\n"
                "Primjer navođenja knjige koja ima jednoga autora:\n"
                "Burckhardt, Jacob. 1953. Eseji iz grčke prošlosti. Zora. Zagreb.\n\n"
                "Primjer navođenja knjige koja ima do tri autora:\n"
                "Silić, Josip; Pranjković, Ivo. 2007. Gramatika hrvatskoga jezika. Školska knjiga. Zagreb."
            )
        doc.add_page_break()

    # Spremanje u memoriju (umjesto na disk) kako bi se moglo preuzeti preko weba
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- WEB SUČELJE ---
st.title("Generator školskog referata")
st.write("Ispunite podatke ispod, a sustav će automatski generirati formatiran Word dokument.")

# Polja za unos
naslov = st.text_input("Naslov referata")
predmet = st.text_input("Naziv predmeta (npr. povijesti)")
mentor = st.text_input("Ime i prezime mentora (s titulom)")
ucenici = st.text_input("Ime i prezime učenika")
razred = st.text_input("Razred (npr. 3.a)")
datum = st.text_input("Datum (npr. 3. rujna 2018.)")

# Gumb za generiranje
if st.button("Pripremi dokument"):
    if naslov and predmet and mentor and ucenici and razred and datum:
        word_file = kreiraj_referat(naslov, predmet, mentor, ucenici, razred, datum)
        ime_datoteke = f"Referat_{ucenici.replace(' ', '_')}.docx"
        
        st.success("Dokument je spreman!")
        st.download_button(
            label="⬇️ Preuzmi Referat",
            data=word_file,
            file_name=ime_datoteke,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.error("Molimo ispunite sva polja prije generiranja.")