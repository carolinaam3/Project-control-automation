from app.config_loader import load_config
from app.curva_s_updater import update_curve_s
from app.dashboard_updater import update_dashboard
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

    dashboard_output = update_dashboard(config)
    curve_output = update_curve_s(config, dashboard_output)

    print("LDD carregada com sucesso!")
    print(f"Dashboard atualizado: {dashboard_output}")
    print(f"Curva S atualizada: {curve_output}")


if __name__ == "__main__":
    main()