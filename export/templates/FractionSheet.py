"""
This module contains the class FractionSheet, which is used to export the data of the time fractions.
"""
from export.config import SheetNamesFraction
from export.templates.AbstractSheet import AbstractSheet


class FractionSheet(AbstractSheet):
    """
    This class contains fields that correspond to the column indices in the template. It also contains all the
    properties and methods that AbstractSheet has.
    """
    def __init__(self, template_path):
        super().__init__(template_path, 'fraction 1 of 16', SheetNamesFraction, 3, 7)

    s1_name: int
    s2_name: int
    range: int
    f_hz_s1: int
    measured_s1: int
    f_hz_s2: int
    measured_s2: int
    f_hz_sd1: int
    measured_sd1: int
    f_hz_sd2: int
    measured_sd2: int
    ad: int
    psi: int
    psi_1: int
    s1_p: int
    s2_p: int
