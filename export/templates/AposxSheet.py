from export.config import SheetNamesAposx
from export.templates.AbstractSheet import AbstractSheet


class AposxSheet(AbstractSheet):
    def __init__(self, template_path):
        super().__init__(template_path, 'Apos(x)', SheetNamesAposx, 2, 3)

    xyz_position: int
    avg_ad: int
    max_ad: int
    min_ad: int
    avg_psi: int
    max_psi: int
    min_psi: int
