import pandas as pd
from openpyxl.reader.excel import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from export.config import SheetNamesFt
from export.templates.AbstractSheet import AbstractSheet


class FtSheet(AbstractSheet):
    """
    Class that manages finding column indices in the template. It reimplements the AbstractSheet class, because
    the template for the "ft" sheet is different from the other templates. It is dependent on microphone count
    and existence of some columns are dependent on that.
    """
    def __init__(self, template_path, mono=False):
        xls = pd.ExcelFile(template_path)
        df = pd.read_excel(xls, 'ft', header=None)

        # Find the column indices in the specific sheet of the template
        if mono:
            # Mono
            self.t_fraction = 1
            self.f_range = 2
            self.avg_ad = 3
            self.avg_psi = 4
        else:
            # Stereo
            for enum in SheetNamesFt:
                self.__setattr__(enum.name, self._find_attribute_in_excel(df.values, enum.value))

        self._wb = load_workbook(template_path)
        self.sheet_name = 'ft'
        self.sheet: Worksheet = self._wb['ft']
        self.start_row = 2
        self.template_rows = 16

    # Common for mono and stereo
    t_fraction: int
    f_range: int
    avg_ad: int
    avg_psi: int

    # Only exist for stereo
    max_ad: int
    min_ad: int
    max_psi: int
    min_psi: int
