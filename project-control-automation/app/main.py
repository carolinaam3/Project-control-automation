from app.config_loader import load_config
from app.ldd_processor import load_ldd
from app.validators import validate_required_columns


def main():
    config = load_config("projeto_exemplo")
    df = load_ldd(config)

    missing_columns = validate_required_columns(df, config)

    if missing_columns:
        print("Colunas obrigatórias ausentes:")
        for col in missing_columns:
            print(f"- {col}")
        return

    print("LDD carregada e validada com sucesso!")
    print(df.head())


if __name__ == "__main__":
    main()