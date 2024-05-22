from export.config import SheetNamesAt
from export.templates.AbstractSheet import AbstractSheet


class AtSheet(AbstractSheet):
    def __init__(self, template_path):
        super().__init__(template_path, 'At', SheetNamesAt, 2, 16)

    # Common for mono and stereo
    t_fraction: int
    xyz_position: int
    s1: int
    s2: int
    ad: int
    psi_a: int

    # Only exist for surround
    avg_ad: int
    avg_psi: int
    max_ad: int
    min_ad: int
    max_psi: int
    min_psi: int
