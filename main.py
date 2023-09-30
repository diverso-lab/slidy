from PPTXGenerator import PPTXGenerator

if __name__ == "__main__":
    markdown_file = "ejemplo_presentacion.md"
    template_file = "plantilla.pptx"
    generator = PPTXGenerator(template_file)
    generator.generate_from_markdown(markdown_file)