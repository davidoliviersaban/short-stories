import argparse
from docx import Document

def convert_to_md(input_docx, output_md):
    # Ouvrir le document Word
    doc = Document(input_docx)
    
    # Ouvrir le fichier de sortie en mode écriture
    with open(output_md, 'w', encoding='utf-8') as md_file:
        # Parcourir chaque paragraphe dans le document Word
        for paragraph in doc.paragraphs:
            # Vérifier si le paragraphe est un heading
            if paragraph.style.name.startswith('Heading'):
                # Récupérer le niveau de titre (1 à 9)
                heading_level = int(paragraph.style.name.split()[-1])
                # Ajouter le titre avec le niveau approprié dans le fichier Markdown
                md_file.write('#' * heading_level + ' ' + paragraph.text + '\n\n')
            else:
                # Écrire le texte du paragraphe avec les mises en forme appropriées dans le fichier Markdown
                # Vérifier si le paragraphe fait partie d'une liste à puces
                if 'list' in paragraph.style.name.lower() and paragraph.text.strip() != '':
                    # print (paragraph.style)
                    if 'bullet' in paragraph.style.name.lower():
                        md_file.write('* ')
                    # Vérifier si le paragraphe fait partie d'une liste numérotée
                    else:
                        md_file.write('1. ')
                # Parcourir chaque run dans le paragraphe
                md_text = ''
                for run in paragraph.runs:
                    if run.bold and run.italic:
                        md_text += '***' + run.text.strip() + '*** '
                    # Vérifier si le texte est en gras
                    elif run.bold:
                        md_text += '**' + run.text.strip() + '** '
                    # Vérifier si le texte est en italique
                    elif run.italic:
                        md_text += '*' + run.text.strip() + '* '
                    # Vérifier si le texte est surligné
                    elif run.font.highlight_color is not None:
                        md_text += '> ' + run.text.strip() + '\n'
                    else:
                        md_text += run.text
                
                md_file.write(md_text + '\n\n')

if __name__ == '__main__':
    # Créer un objet ArgumentParser
    parser = argparse.ArgumentParser(description='Convert Word document to Markdown.')

    # Ajouter des arguments
    parser.add_argument('input_docx', help='Input Word document file')
    parser.add_argument('output_md', help='Output Markdown file')

    # Analyser les arguments de la ligne de commande
    args = parser.parse_args()

    # Convertir le document Word en fichier Markdown
    convert_to_md(args.input_docx, args.output_md)
