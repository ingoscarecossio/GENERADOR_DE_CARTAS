
# -*- coding: utf-8 -*-
"""
CLI para generación por lotes (filtrable por ACTOR, newest-first)
Uso:
  python cli.py --excel BASE.xlsx --template MODELO.docx --out outdir --group ACTOR --actors "A;B;C"
"""
import argparse, os, json, pandas as pd
from core.backend import guess_mapping, prepare_dataframe
from core.funcionalidades import generate_letters_per_group

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--excel", required=True)
    p.add_argument("--template", required=True)
    p.add_argument("--out", required=True)
    p.add_argument("--group", default="ACTOR")
    p.add_argument("--actors", default=None, help="Lista separada por ';' para filtrar ACTOR")
    p.add_argument("--newest-first", action="store_true")
    args = p.parse_args()

    os.makedirs(args.out, exist_ok=True)
    df = pd.read_excel(args.excel)

    mapping = guess_mapping(df)
    if args.group != "ACTOR":
        mapping["grupo"] = args.group
    if not all(mapping.get(k) for k in ("actor","mesa","nivel","fecha","dato")):
        raise SystemExit("No se detectaron todas las columnas requeridas. Renombre o use columnas correctas.")

    work = prepare_dataframe(df, mapping)
    if args.actors:
        allowed = set([x.strip() for x in args.actors.split(";") if x.strip()])
        work = work[work["ACTOR"].astype(str).isin(allowed)]

    with open(args.template, "rb") as f:
        tpl = f.read()
    outputs, errors, index_df = generate_letters_per_group(
        work, tpl,
        group_field=("GRUPO" if "GRUPO" in work.columns and args.group != "ACTOR" else "ACTOR"),
        placeholders_per_group=None,
        table_index=None,
        newest_first=args.newest_first
    )

    for fname, data in outputs.items():
        with open(os.path.join(args.out, fname), "wb") as w:
            w.write(data)
    index_df.to_excel(os.path.join(args.out, "indice_cartas.xlsx"), index=False)
    if errors:
        print(json.dumps(errors, ensure_ascii=False, indent=2))
    print(f"Generados: {len(outputs)}")

if __name__ == "__main__":
    main()
