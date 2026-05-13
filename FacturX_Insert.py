import sys
from facturx import generate_from_file


def insert_facturx(xml_file, pdf_file):
    generate_from_file(
        pdf_file,
        xml_file,
        flavor="factur-x",
        level="en16931",
        output_pdf_file=pdf_file
    )


def main():
    if len(sys.argv) != 3:
        print("Usage : python insert_facturx.py fichier.xml fichier.pdf")
        sys.exit(1)

    xml_file = sys.argv[1]
    pdf_file = sys.argv[2]

    try:
        insert_facturx(xml_file, pdf_file)
        print("Factur-X intégré avec succès.")
        sys.exit(0)

    except Exception as e:
        print(f"Erreur : {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()