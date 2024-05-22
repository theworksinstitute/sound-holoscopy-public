"""
This module tests integration of "ft" template. It verifies that columns are found correctly, cells are copied
correctly, styles match, etc.
"""
from export.config import template_path, resources_dir
from export.templates.FtSheet import FtSheet


def test_template_ft_stereo():
    template = FtSheet(template_path, mono=False)
    assert template.t_fraction == 1
    assert template.f_range == 2
    assert template.avg_ad == 3
    assert template.max_ad == 4
    assert template.min_ad == 5
    assert template.avg_psi == 6
    assert template.max_psi == 7
    assert template.min_psi == 8


def test_template_ft_mono():
    template = FtSheet(template_path, mono=True)
    assert template.t_fraction == 1
    assert template.f_range == 2
    assert template.avg_ad == 3
    assert template.avg_psi == 4
    assert hasattr(template, 'max_ad') is False
    assert hasattr(template, 'min_ad') is False
    assert hasattr(template, 'max_psi') is False
    assert hasattr(template, 'min_psi') is False


def test_copy_template_row():
    template = FtSheet(template_path, mono=False)
    template.copy_template_values(template.sheet, 18, 1, styles_only=True)
    assert template.sheet['H18']._style == template.sheet['H17']._style


def test_copy_template_rows():
    template = FtSheet(template_path, mono=False)
    template.copy_template_values(template.sheet, 18, 16, styles_only=True)
    assert template.sheet['H17']._style == template.sheet['H33']._style


def test_template_starting_row():
    template = FtSheet(template_path, mono=False)
    starting_row = template.start_row
    assert template.sheet[starting_row - 1][0].value == 't fraction'
    assert template.sheet[starting_row - 1][1].value == 'f range'
    assert template.sheet[starting_row - 1][2].value == 'average AD (-dB)'
    assert template.sheet[starting_row - 1][5].value == 'average 𝜓 (degrees)'
    assert template.sheet[starting_row][0].value == 't1'
    assert template.sheet[starting_row + 1][0].value == 't2'
    assert template.sheet[starting_row][1].value == '50-52'
