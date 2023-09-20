# Conversor Markdown a PowerPoint

Este script convierte un archivo en formato Markdown a una presentación de PowerPoint utilizando una plantilla predefinida.

## Requisitos

Para ejecutar este script, es necesario tener instalados los siguientes paquetes:

- `pptx`: para trabajar con archivos de PowerPoint.
- `markdown`: para procesar archivos en formato Markdown.
- `bs4` (BeautifulSoup): para procesar el HTML generado por el parser de Markdown.

Puedes instalarlos con pip:

```bash
pip install python-pptx markdown beautifulsoup4
```

También se pueden instalar todos a la vez:

```bash
pip install -r requirements.txt
```

## Uso

Para utilizar el script:

1. Asegúrate de que el archivo `ejemplo_presentacion.md` y la plantilla `plantilla.pptx` se encuentren en el mismo directorio que el script.
2. Ejecuta el script:

```bash
python main.py
```

Esto generará un archivo `output.pptx` con la presentación.

## Personalización

- **Template**: Puedes cambiar la plantilla de PowerPoint que se utiliza editando la variable `template_file` en la sección `if __name__ == "__main__":`.
- **Archivo Markdown**: Puedes cambiar el archivo Markdown que se convierte editando la variable `markdown_file` en la misma sección.

## Funcionamiento

El script toma el archivo en formato Markdown y busca títulos y subtítulos, los cuales utiliza para generar el contenido de las diapositivas en la presentación de PowerPoint. 

Utiliza una plantilla de PowerPoint predefinida, donde se espera que ciertos placeholders, como `#main_title`, estén presentes para ser reemplazados con el contenido apropiado del archivo Markdown.

## Actualizar dependencias

Para actualizar todas las dependencias del proyecto automáticamente, ejecute:

```
chmod +x update_dependencies.sh && ./update_dependencies.sh
```

Nota: es responsabilidad del desarrollador comprobar que la actualización de dependencias no ha roto ninguna funcionalidad y cada dependencia mantiene la compatibilidad con versiones anteriores. ¡Utilice el script con cuidado!

Traducción realizada con la versión gratuita del traductor www.DeepL.com/Translator
## Contribuciones

Si encuentras algún problema o tienes sugerencias de mejoras, no dudes en abrir un Issue o Pull Request en este repositorio.
