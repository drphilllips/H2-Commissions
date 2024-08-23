
from enum import Enum

class FileLoc(Enum):
    BASE = "C:/Users/dylan/Desktop/H2 Commissions/"
    BACKUP = BASE + "Backup/"
    LOOKUP = BASE + "Lookup/"
    OUTPUT = BASE + "Output/"
    INPUT = BASE + "Input/"
    MASTER = BASE + "H2 Commissions Master.xlsx"
    ROOT_CUSTOMER_DICTIONARY = LOOKUP + "Root Customer Dictionary.xlsx"
    FSE_DICTIONARY = LOOKUP + "H2 ALL TERRITORIES ACCOUNT LIST.xlsx"

