def color_print(text, color):
    """
    can print color text
    """
    print(f"\x1b[38;5;{color}m{text}\x1b[m")
