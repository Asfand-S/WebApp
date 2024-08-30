import io
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.units import cm
from reportlab.lib import colors

# Get styles
styles = getSampleStyleSheet()
style             = ParagraphStyle('MyStyle', styles['Normal'], fontName="Helvetica", fontSize=16, leftIndent=0.5*cm, spaceAfter=0.5*cm, leading=20)
margin_style      = ParagraphStyle('MyStyle', styles['Normal'], fontName="Helvetica", fontSize=16, leftIndent=0.5*cm, spaceAfter=0.5*cm, leading=20, rightIndent=6*cm)
centre_style      = ParagraphStyle('MyStyle', styles['Normal'], fontName="Helvetica", fontSize=20, leftIndent=2*cm, rightIndent=2*cm, alignment=1, leading=25)
centre_bold_style = ParagraphStyle('MyStyle', styles['Normal'], fontName="Helvetica-Bold", fontSize=20, leftIndent=2*cm, rightIndent=2*cm, alignment=1, leading=25)

def calculate_age(birthdate):
    try:
        # Ensure the birthdate is a datetime object
        birthdate = str(birthdate)
        birthdate = datetime.strptime(birthdate, "%Y-%m-%d %H:%M:%S")

        today = datetime.today()
        age = today.year - birthdate.year

        # Adjust for the case where the birthday has not yet occurred this year
        if (today.month, today.day) < (birthdate.month, birthdate.day):
            age -= 1

        return age
    except:
        return "Pas d'informations"

# Function to draw the yellow border on each page
def add_yellow_border(canvas, doc, photo_link="www.emploicesu.fr", cv_link="www.emploicesu.fr"):
    # Set the border color to yellow
    canvas.setStrokeColorRGB(255/256, 222/256, 88/256)
    canvas.setFillColorRGB(255/256, 222/256, 88/256)
    
    # Draw the yellow border (1.5 cm thick) around the page
    canvas.rect(0, 0, doc.width + 3*cm, doc.height + 3*cm, stroke=1, fill=1)
    
    # Set the fill color back to white for the rest of the content area
    canvas.setFillColor(colors.white)
    canvas.rect(1.5*cm, 1.5*cm, doc.width, doc.height, stroke=0, fill=1)

    canvas.setFont("Helvetica", 20)
    canvas.setFillColor(colors.black)
    canvas.drawString(7.4*cm, 0.7*cm, "www.emploicesu.fr")

    canvas.drawImage("./static/photo_icon.PNG", 13.5*cm, 22*cm, 5*cm, 5*cm)
    canvas.linkURL(photo_link, (13.5*cm, 22*cm, 18.5*cm, 27*cm), relative=1)

    canvas.drawImage("./static/cv_icon.PNG", 13.5*cm, 3*cm, 5*cm, 5*cm)
    canvas.linkURL(cv_link, (13.5*cm, 3*cm, 18.5*cm, 8*cm), relative=1)

def on_later_pages(canvas, doc):
    # Set the border color to yellow
    canvas.setStrokeColorRGB(255/256, 222/256, 88/256)
    canvas.setFillColorRGB(255/256, 222/256, 88/256)
    
    # Draw the yellow border (1.5 cm thick) around the page
    canvas.rect(0, 0, doc.width + 3*cm, doc.height + 3*cm, stroke=1, fill=1)
    
    # Set the fill color back to white for the rest of the content area
    canvas.setFillColor(colors.white)
    canvas.rect(1.5*cm, 1.5*cm, doc.width, doc.height, stroke=0, fill=1)

    canvas.setFont("Helvetica", 20)
    canvas.setFillColor(colors.black)
    canvas.drawString(7.4*cm, 0.7*cm, "www.emploicesu.fr")

# Function to create a PDF for each row
def create_pdf(row, index, save_folder):
    try:
        step = 1
        items = ["Prénom", "Nom", "Numéro de téléphone", "Quel âge avez-vous ? ", "Ville", "Code postal", "Nationalité ", "Permis B+ véhicule", 
                "Combien d'années d'expériences avez-vous ?", "Précisez l'intitulé de votre diplôme?", "Parlez-vous français", 
                "Avez-vous des références employeur à communiquer sur demande ?: ", 
                "Souhaitez-vous réaliser des formations complémentaires ? ", 
                "Quelle est votre fourchette de salaire minimum net en CESU ?", "Quels sont vos centres d'intérêts ?", "Matin", 
                "Après-midi", "Soirées", "Vos disponibilités varient-elles régulièrement ?"]
        data = []
        for item in items:
            value = str(row[item])
            if value == "":
                data.append("Pas d'informations")
            else:
                data.append(value)

        step = 2
        elements = [Spacer(20*cm, 0.5*cm)]

        paragraph = Paragraph(data[0], style)
        elements.append(paragraph)
        paragraph = Paragraph(data[1], style)
        elements.append(paragraph)
        paragraph = Paragraph(f"+{data[2]}", style)
        elements.append(paragraph)
        paragraph = Paragraph(f"{calculate_age(row["Quel âge avez-vous ? "])}", style)
        elements.append(paragraph)
        paragraph = Paragraph(data[4], style)
        elements.append(paragraph)
        paragraph = Paragraph(data[5], style)
        elements.append(paragraph)
        paragraph = Paragraph(data[6], style)
        elements.append(paragraph)
        paragraph = Paragraph(f"Permis B+ Véhicule: {data[7]}", style)
        elements.append(paragraph)
        paragraph = Paragraph(f"<u>Compétences</u>", style)
        elements.append(paragraph)
        paragraph = Paragraph(f"Niveau d’expérience: {data[8]}", style)
        elements.append(paragraph)
        paragraph = Paragraph(f"Diplômé(e) du secteur: {data[9]}", style)
        elements.append(paragraph)
        paragraph = Paragraph(f"Parle français: {data[10]}", style)
        elements.append(paragraph)
        paragraph = Paragraph(f"Avez-vous des références et pourriez-vous les communiquer sur demande? {data[11]}", style)
        elements.append(paragraph)
        paragraph = Paragraph(f"Seriez-vous prêt(e) à faire des formationscomplémentaires financées par l’Etat (Iperia.fr)? {data[12]}", style)
        elements.append(paragraph)
        paragraph = Paragraph(f"Vous pouvez cliquer sur le CV ci-dessous pour le télécharger: ", margin_style)
        elements.append(paragraph)
        paragraph = Paragraph(f"P ̈rétentions salariales, à partir de: {data[13]}", margin_style)
        elements.append(paragraph)
        elements.append(PageBreak())
        paragraph = Paragraph(f"Les centres d’intérêts: {data[14]}", style)
        elements.append(paragraph)
        paragraph = Paragraph(f"<u>Disponibilités Générales soumises à évolution</u>", style)
        elements.append(paragraph)
        paragraph = Paragraph(f"Matin: {data[15]}", style)
        elements.append(paragraph)
        paragraph = Paragraph(f"Après-midi: {data[16]}", style)
        elements.append(paragraph)
        paragraph = Paragraph(f"Soirées: {data[17]}", style)
        elements.append(paragraph)
        paragraph = Paragraph(f"Mes disponibilités peuvent-elles être soumises àvariation? {data[18]}", style)
        elements.append(paragraph)
        elements.append(Spacer(20*cm, 2*cm))
        paragraph = Paragraph(f"Si vous aimez votre profil, n’hésitez pas àpartager le lien au plus grand nombre", centre_style)
        elements.append(paragraph)
        elements.append(Spacer(20*cm, 1*cm))
        link = "www.emploicesu.fr"
        address = '<link href="' + 'www.emploicesu.fr' + '"><u>' + link + '</u></link>'
        paragraph = Paragraph(address, centre_bold_style)
        elements.append(paragraph)
        paragraph = Paragraph(f"Là pour tous vos besoins !", centre_bold_style)
        elements.append(paragraph)
        elements.append(Image("./static/heart_icon.PNG", width=3*cm, height=5*cm))

        step = 3
        # Build the PDF with a custom page template to add the yellow border
        cv_link = str(row["Téléchargez votre CV en ligne en vérifiant que vos informations soient bien à jour "])
        photo_link = str(row["Ajoutez une Photo d'identité ou photo autre professionnelle"])

        output = io.BytesIO()
        doc = SimpleDocTemplate(output, pagesize=A4,
                                leftMargin=1.5*cm, rightMargin=1.5*cm,
                                topMargin=1.5*cm, bottomMargin=1.5*cm)

        doc.build(elements, onFirstPage=lambda canvas, doc: add_yellow_border(canvas, doc, photo_link, cv_link), onLaterPages=on_later_pages)
        output.seek(0)
        return output
    except Exception as e:
        return (index, e, step)

def generate_pdf(df, save_folder):
    if save_folder:
        df.fillna("", inplace=True)
        # Generate a PDF for each row
        results = []
        for index, row in df.iterrows():
            # if index == 25:
            #     create_pdf(row, index)
            #     break
            result = create_pdf(row, index, save_folder)
            return result
            if result != True:
                results.append(result)

        if results:
            print(f'Total Files: {index+1}\nFailed files: {len(results)}\nFailed at rows: {", ".join([str(r[0]) for r in results])}')
        else:
            print(f'File successfully saved to {save_folder}')

