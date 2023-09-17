from pptx import Presentation
import markdown
from bs4 import BeautifulSoup
from copy import deepcopy


def duplicate_slide(prs, slide_index):
    """Duplica una diapositiva en una presentación."""
    source = prs.slides[slide_index]
    slide_layout = source.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)

    for shape in reversed(source.shapes):
        el = shape.element
        new_el = deepcopy(el)
        new_slide.shapes._spTree.insert(2, new_el)

    return new_slide


def replace_text_while_keeping_format(shape, placeholder, new_text):
    """Reemplaza texto en una forma mientras mantiene el formato original."""
    if placeholder in shape.text:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, new_text)


def markdown_to_pptx_with_template(input_file, template_path):
    # Lectura del archivo markdown y transformación a HTML
    with open(input_file, 'r') as f:
        content = f.read()
    html_content = markdown.markdown(content)
    soup = BeautifulSoup(html_content, 'html.parser')

    # Extracción de información de cabecera
    title = soup.find('h1').get_text()
    authors_line = soup.find('h2').get_text()
    workshop = soup.find('h3').get_text()

    # Carga de la plantilla en Python
    prs = Presentation(template_path)

    # Rellenar todas las diapositivas con #main_title, #authors, y #workshop
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                replace_text_while_keeping_format(shape, "#main_title", title)
                replace_text_while_keeping_format(shape, "#authors", authors_line.replace("Authors:", "").strip())
                replace_text_while_keeping_format(shape, "#workshop", workshop)

    # Primero, calculamos el total de diapositivas de contenido
    total_content_slides = sum(1 for section in soup.find_all('h2')[1:] for _ in section.find_all_next('h3'))

    # Crear diapositivas de contenido
    sections = soup.find_all('h2')[1:]  # Excluir el primer h2 que son autores
    section_counter = 1
    content_slide_counter = 0  # Añadimos un contador para las diapositivas de contenido

    for section in sections:
        subtitles = section.find_all_next('h3')
        subtitle_counter = 1
        for subtitle in subtitles:
            slide = duplicate_slide(prs, 1)  # Duplicamos la segunda diapositiva
            content_slide_counter += 1  # Incrementamos el contador de diapositivas de contenido
            for shape in slide.shapes:
                if shape.has_text_frame:
                    section_title = f"{section_counter}. {section.get_text()}"
                    section_subtitle = f"{section_counter}.{subtitle_counter}. {subtitle.get_text()}"
                    slide_number_format = f"{content_slide_counter:02} - {total_content_slides:02}"  # Formato XX - YY

                    replace_text_while_keeping_format(shape, "#section_title", section_title)
                    replace_text_while_keeping_format(shape, "#section_subtitle", section_subtitle)
                    replace_text_while_keeping_format(shape, "#main_title", title)
                    replace_text_while_keeping_format(shape, "#slide_number", slide_number_format)
            subtitle_counter += 1
        section_counter += 1

    # Eliminar el slide base (el segundo slide en este caso)
    slide_id = prs.slides._sldIdLst[1].rId
    prs.part.drop_rel(slide_id)
    del prs.slides._sldIdLst[1]

    # Guardar
    output_path = "output.pptx"
    prs.save(output_path)
    print(f"Presentación guardada en {output_path}")


if __name__ == "__main__":
    markdown_file = "ejemplo_presentacion.md"
    template_file = "plantilla.pptx"
    markdown_to_pptx_with_template(markdown_file, template_file)
