from pathlib import Path

import pandas as pd


def load_ldd(config: dict) -> pd.DataFrame:
    """
    Carrega a LDD usando as configurações do projeto.
    """

    input_folder = Path(config["paths"]["input_folder"])
    ldd_file = config["paths"]["ldd_file"]
    header_row = config["ldd"]["header_row"]

    file_path = input_folder / ldd_file

    if not file_path.exists():
        raise FileNotFoundError(f"LDD não encontrada: {file_path}")

    df = pd.read_excel(file_path, header=header_row)
    return df