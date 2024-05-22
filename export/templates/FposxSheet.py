from export.config import SheetNamesFposx
from export.templates.AbstractSheet import AbstractSheet


class FposxSheet(AbstractSheet):
    def __init__(self, template_path):
        super().__init__(template_path, 'fpos(x)', SheetNamesFposx, 2, 3)

    xyz_position: int
    f_range: int
    avg_ad: int
    max_ad: int
    min_ad: int
    avg_psi: int
    max_psi: int
    min_psi: int
