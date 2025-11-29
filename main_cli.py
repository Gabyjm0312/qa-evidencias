# main_cli.py

from evidencias_core import (
    parse_args,
    generar_evidencias_desde_excel,
)


def main():
    # Leemos los argumentos de línea de comandos (los mismos de antes)
    a = parse_args()

    # Pasamos esos argumentos a la función reutilizable
    generar_evidencias_desde_excel(
        excel_path=a.excel or "",           # si viene None, el core autodetecta
        hoja=a.hoja,
        salida=a.salida,
        sin_consolidado=a.sin_consolidado,
        sin_individuales=a.sin_individuales,
        map_json=a.map_json,
    )


if __name__ == "__main__":
    main()
