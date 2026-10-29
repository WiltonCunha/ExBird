"""Procura em qual tabela/coluna de TEXTO do banco Firebird um valor está gravado.

Uso:
  python localizar_valor.py "0202400" --db C:\\caminho\\banco.fdb [--host localhost] [--port 3050]
                            [--user SYSDBA] [--senha masterkey] [--charset WIN1252] [--tabelas ITENS,ITENS_FISCAL]

Como usar pra achar onde o ERP guarda o CEST:
  1. No ERP, grave um produto de teste com um CEST bem único, ex: 0123456
  2. Rode: python localizar_valor.py 0123456 --db ...
  3. Ele lista as colunas onde esse valor exato aparece (tabela, coluna e exemplos da linha).
Só lê o banco, não altera nada.
"""
import argparse
import sys

import core


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("valor")
    ap.add_argument("--db", required=True)
    ap.add_argument("--host", default="localhost")
    ap.add_argument("--port", default="3050")
    ap.add_argument("--user", default="SYSDBA")
    ap.add_argument("--senha", default="masterkey")
    ap.add_argument("--charset", default="WIN1252")
    ap.add_argument("--tabelas", default="", help="limitar a estas tabelas (separadas por vírgula)")
    a = ap.parse_args()

    conn = core.open_conn(dict(host=a.host, port=a.port, db=a.db, user=a.user,
                               password=a.senha, charset=a.charset))
    cur = conn.cursor()
    cur.execute(
        "SELECT TRIM(rf.RDB$RELATION_NAME), TRIM(rf.RDB$FIELD_NAME) "
        "FROM RDB$RELATION_FIELDS rf "
        "JOIN RDB$FIELDS f ON f.RDB$FIELD_NAME = rf.RDB$FIELD_SOURCE "
        "JOIN RDB$RELATIONS r ON r.RDB$RELATION_NAME = rf.RDB$RELATION_NAME "
        "WHERE COALESCE(r.RDB$SYSTEM_FLAG, 0) = 0 AND f.RDB$FIELD_TYPE IN (14, 37) "
        "ORDER BY 1, rf.RDB$FIELD_POSITION")
    cols = cur.fetchall()
    if a.tabelas.strip():
        so = {t.strip().upper() for t in a.tabelas.split(",")}
        cols = [c for c in cols if c[0].upper() in so]

    print(f'Procurando "{a.valor}" em {len(cols)} colunas de texto...', flush=True)
    achou = 0
    for n, (t, c) in enumerate(cols, 1):
        try:
            cur.execute(f'SELECT FIRST 3 * FROM "{t}" WHERE TRIM("{c}") = ?', (a.valor,))
            linhas = cur.fetchall()
            if linhas:
                achou += 1
                nomes = [d[0] for d in cur.description][:6]
                print(f"\n>>> {t}.{c}  (exemplos de linha, primeiras colunas: {nomes})")
                for l in linhas:
                    print("   ", [str(x)[:25] for x in l[:6]])
        except Exception:
            conn.rollback()  # coluna/tabela que não deu pra consultar: segue
        if n % 200 == 0:
            print(f"  ... {n}/{len(cols)}", flush=True)
    print(f"\nFim. {achou} coluna(s) com o valor." if achou else
          "\nFim. Nenhuma coluna de texto tem esse valor exato.\n"
          "Então o ERP ou guarda em coluna numérica, ou monta o zero na tela a partir de outra informação.")
    conn.close()


if __name__ == "__main__":
    sys.exit(main())
