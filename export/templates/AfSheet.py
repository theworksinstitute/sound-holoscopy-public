from export.config import SheetNamesAf
from export.templates.AbstractSheet import AbstractSheet


class AfSheet(AbstractSheet):
    def __init__(self, template_path):
        super().__init__(template_path, 'Af', SheetNamesAf, 2, 7)

    xyz_position: int
    f_range: int
    avg_ad: int
    max_ad: int
    min_ad: int
    avg_psi: int
    max_psi: int
    min_psi: int
