import pytest

from export.colors import get_colors, get_odd_colors


def test_get_colors():
    color_generator = get_colors()
    assert next(color_generator) == "0000ff"  # Deep Blue (1st octave)

    assert next(color_generator) == "0000ff"  # Deep Blue (2nd octave)
    assert next(color_generator) == "ffff00"  # Yellow (2st octave)

    assert next(color_generator) == "0000ff"  # Deep Blue (3rd octave)
    assert next(color_generator) == "ff007f"  # Red (3rd octave)
    assert next(color_generator) == "ffff00"  # Yellow (3rd octave)
    assert next(color_generator) == "00ff7f"  # Green (3rd octave)

    assert next(color_generator) == "0000ff"  # Deep Blue (4th octave)
    assert next(color_generator) == "bf00ff"  # Pink (4th octave)
    assert next(color_generator) == "ff007f"  # Red (4th octave)
    assert next(color_generator) == "ff3f00"  # Orange (4th octave)
    assert next(color_generator) == "ffff00"  # Yellow (4th octave)
    assert next(color_generator) == "3fff00"  # Light Green (4th octave)
    assert next(color_generator) == "00ff7f"  # Green (4th octave)
    assert next(color_generator) == "00bfff"  # Light Blue (4th octave)

    assert next(color_generator) == "0000ff"  # Deep Blue (1st octave)
    # ... and so on.


def test_get_odd_numbers():
    color_generator = get_odd_colors()
    assert next(color_generator) == "0000ff"  # Deep Blue (1st octave)
    assert next(color_generator) == "ffff00"  # Yellow (2st octave)
    assert next(color_generator) == "ff007f"  # Red (3rd octave)
    assert next(color_generator) == "00ff7f"  # Green (3rd octave)
    assert next(color_generator) == "bf00ff"  # Pink (4th octave)
    assert next(color_generator) == "ff3f00"  # Orange (4th octave)
    assert next(color_generator) == "3fff00"  # Light Green (4th octave)
    assert next(color_generator) == "00bfff"  # Light Blue (4th octave)
    # ... and so on.
