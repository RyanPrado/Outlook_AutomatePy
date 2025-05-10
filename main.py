import pandas as pd
import sys
import time
import win32com.client as win32
from utils.parser import parser
from utils.logging import logging
from pathlib import Path

logger = logging.getLogger(__name__)
args = parser.parse_args()


def open_excel_dataframe(path_str: str) -> pd.DataFrame:
    """Open the excel file and return to the first sheet

    Args:
        path_str (str): Excel file path

    Returns:
        pd.DataFrame: First sheet for excel workbook
    """
    try:
        return pd.read_excel(path_str, sheet_name=0)
    except FileNotFoundError:
        logger.error("O arquivo fornecido está corrompido ou não existe!")
    except ValueError:
        logger.error("O arquivo fornecido não está em um formatado suportado!")
    except Exception as e:
        logger.error(f"Ocorreu um erro:\n{e}")


def main():
    if sys.platform != "win32":
        logger.info("Este código só pode ser executado em um ambiente Windows.")
        sys.exit()

    try:
        olApp = win32.Dispatch("Outlook.Application")
        olNS = olApp.GetNameSpace("MAPI")
    except Exception as e:
        logger.error(
            "Verifique se você possui o Outlook Microsoft Office 365, instalado em sua maquina."
        )

    sender: str = args.sender
    logger.debug("Trying read excel file...")
    df = open_excel_dataframe(args.excel)
    logger.debug("Opened excel file!")
    logger.debug("Looping in rows...")
    for _, v in df.iterrows():
        emails_target: tuple[str] = str(v.iloc[0]).split(";")
        subject: str = v.iloc[1]
        body: str = v.iloc[2]
        attachments_str = str(v.iloc[3])
        attachments_files: list[Path] = []
        if (
            attachments_str != ""
            and attachments_str != "nan"
            and attachments_str != "None"
        ):
            attachments_path: Path = Path(str(attachments_str))
            attachments_files = [
                x for x in attachments_path.iterdir() if not x.is_dir()
            ]

        for email_target in emails_target:
            mail = olApp.CreateItem(0)
            mail.To = str(email_target)
            mail.Subject = subject
            mail.BodyFormat = 1
            mail.Body = body
            mail.Sender = sender
            for file in attachments_files:
                mail.Attachments.Add(rf"{str(file)}")

            logger.info(
                f"Email enviado para {mail.To}, com {len(attachments_files)} anexo(s)"
            )
            mail.Send()
            time.sleep(0.5)
        time.sleep(1)


if __name__ == "__main__":
    main()
