from pathlib import Path
import yaml


def load_config(config_name: str) -> dict:
    """
    Carrega um arquivo de configuração YAML.

    Args:
        config_name: Nome do arquivo sem a extensão.
                     Exemplo: "projeto_exemplo"

    Returns:
        Dicionário contendo todas as configurações do projeto.
    """

    config_path = Path("config") / f"{config_name}.yaml"

    if not config_path.exists():
        raise FileNotFoundError(
            f"Arquivo de configuração não encontrado: {config_path}"
        )

    with open(config_path, "r", encoding="utf-8") as file:
        config = yaml.safe_load(file)

    return config