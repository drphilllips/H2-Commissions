
from enum import Enum

class FileLoc(Enum):
    BASE = "../../"
    BACKUP = BASE + "Backup/"
    LOOKUP = BASE + "Lookup/"
    OUTPUT = BASE + "Output/"
    INPUT = BASE + "Input/"
    MASTER = BASE + "H2 Commissions Master.xlsx"
    FIELD_MAPPINGS = LOOKUP + "Field Mappings.xlsx"
    LOOKUP_MATRIX = LOOKUP + "Lookup Matrix.xlsx"
    FORMAT_MATRIX = LOOKUP + "Format Matrix.xlsx"

