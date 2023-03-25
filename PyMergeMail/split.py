import pandas as pd
from math import remainder

async def split(data_frame, divider):
    df = pd.DataFrame(data_frame)
    rows = df.shape[0]
    # remainder = remainder(rows, divider)
    return {rows/divider : remainder(rows, divider)}
