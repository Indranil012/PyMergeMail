def color_print(text: str, colour: int):
    """
    can print color text
    """
    print(f"\x1b[38;5;{colour}m{text}\x1b[m")

def color(text: str, colour: int) -> str:
    return f"\x1b[38;5;{colour}m{text}\x1b[m"
