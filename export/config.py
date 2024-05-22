from os import path
from config import ROOT_DIR
from enum import Enum

resources_dir = path.join(ROOT_DIR, 'resources')
sheets_dir = path.join(resources_dir, 'sheets')
sounds_dir = path.join(resources_dir, 'sounds')
template_path = path.join(sheets_dir, 'template.xlsx')


class SheetNamesFraction(Enum):
    """
    Enum that stores default column values for finding the right column in the fraction sheet of the template.
    """
    s1_name = "s1"
    s2_name = "s2"
    range = "f range"
    f_hz_s1 = "f (Hz) s1"
    measured_s1 = "s1 MEASURED (-dB)"
    f_hz_s2 = "f (Hz) s2"
    measured_s2 = "s2 MEASURED (-dB)"
    f_hz_sd1 = "f (Hz) sD1"
    measured_sd1 = "sD1 MEASURED (-dB)"
    f_hz_sd2 = "f (Hz) sD2"
    measured_sd2 = "sD2 MEASURED (-dB)"
    ad = "AD (-dB)"
    psi = "𝜓 (degrees)"
    psi_1 = "𝜓1 (degrees)"
    s1_p = "s1 (p)"
    s2_p = "s2 (p)"


class SheetNamesFt(Enum):
    """
    Enum that stores default column values for finding the right column in the "ft" sheet of the template.
    """
    t_fraction = "t fraction"
    f_range = "f range"
    avg_ad = "average AD (-dB)"
    max_ad = "max AD (-dB)"
    min_ad = "min AD (-dB)"
    avg_psi = "average 𝜓 (degrees)"
    max_psi = "max 𝜓 (degrees)"
    min_psi = "min 𝜓 (degrees)"


class SheetNamesFposx(Enum):
    """
    Enum that stores default column values for finding the right column in the "fpos(x)" sheet of the template.
    """
    xyz_position = "xyz position"
    f_range = "f range"
    avg_ad = "average AD (-dB)"
    max_ad = "max AD (-dB)"
    min_ad = "min AD (-dB)"
    avg_psi = "average 𝜓 (degrees)"
    max_psi = "max 𝜓 (degrees)"
    min_psi = "min 𝜓 (degrees)"


class SheetNamesAf(Enum):
    """
    Enum that stores default column values for finding the right column in the "fpos(x)" sheet of the template.
    """
    xyz_position = "xyz position"
    f_range = "f range"
    avg_ad = "average AD (-dB)"
    max_ad = "max AD (-dB)"
    min_ad = "min AD (-dB)"
    avg_psi = "average 𝜓 (degrees)"
    max_psi = "max 𝜓 (degrees)"
    min_psi = "min 𝜓 (degrees)"


class SheetNamesAt(Enum):
    """
    Enum that stores default column values for finding the right column in the "At" sheet of the template.
    """
    t_fraction = "t fraction"
    xyz_position = "xyz position"
    s1 = "s1 (dBA)"
    s2 = "s2 (dBA)"
    ad = "AD (dBA)"
    psi_a = "𝜓A (degrees)"
    avg_ad = "average AD (-dBA)"
    avg_psi = "average 𝜓A (degrees)"
    max_ad = "max AD (-dBA)"
    min_ad = "min AD (-dBA)"
    max_psi = "max 𝜓A (degrees)"
    min_psi = "min 𝜓A (degrees)"



class SheetNamesAposx(Enum):
    """
    Enum that stores default column values for finding the right column in the "Apos(x)" sheet of the template.
    """
    xyz_position = "xyz position"
    avg_ad = "average AD (-dBA)"
    max_ad = "max AD (-dBA)"
    min_ad = "min AD (-dBA)"
    avg_psi = "average 𝜓A (degrees)"
    max_psi = "max 𝜓a (degrees)"
    min_psi = "min 𝜓A (degrees)"
