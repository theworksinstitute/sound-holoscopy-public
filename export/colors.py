import colorsys

starting_color = 240


# https://github.com/theworksinstitute/sound-holoscopy/issues/61
# https://en.wikipedia.org/wiki/HSL_and_HSV#/media/File:Hsl-hsv_models.svg
def get_colors():
    i = 0
    while True:
        octave_count = 2 ** i
        rotation_angle = 360 / octave_count
        for j in range(octave_count):
            angle = (starting_color + j * rotation_angle) % 360
            yield rgb_to_hex(colorsys.hsv_to_rgb(angle / 360, 1, 1))
        i += 1


def get_odd_colors():
    count = 0
    color_generator = get_colors()

    for color in color_generator:
        if count % 2 == 0:
            yield color
        count += 1


def rgb_to_hex(rgb):
    return ''.join(f'{int(c * 255):02x}' for c in rgb)
