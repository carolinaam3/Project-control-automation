def validate_required_columns(df, config):
    """
    Verifica se as colunas obrigatórias existem na LDD.
    Retorna uma lista com as colunas ausentes.
    """

    required_columns = [
        config["ldd_columns"]["item"],
        config["ldd_columns"]["discipline"],
        config["ldd_columns"]["document_number"],
        config["ldd_columns"]["client_document"],
        config["ldd_columns"]["status"],
        config["ldd_columns"]["planned_issue_date"],
        config["ldd_columns"]["actual_issue_date"],
    ]

    missing_columns = [col for col in required_columns if col not in df.columns]

    return missing_columns