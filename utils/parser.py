import argparse

parser = argparse.ArgumentParser(
    prog="Outlook Automate",
    description="Automizador de envio de emails através do Outlook Office 365",
)
parser.add_argument("excel")
parser.add_argument("sender")
