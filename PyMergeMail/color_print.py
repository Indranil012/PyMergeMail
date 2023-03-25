def color_print(text, colour):
    """
    can print color text
    """
    print(f"\x1b[38;5;{colour}m{text}\x1b[m")

def color(text, colour):
    return f"\x1b[38;5;{colour}m{text}\x1b[m"

