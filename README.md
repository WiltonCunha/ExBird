# ExBird - Conversor

Criado por [Wilton Cunha](https://www.linkedin.com/in/wiltonendril/).

Atualiza colunas de uma tabela Firebird com os valores de uma planilha Excel, usando uma coluna de referência (ex: ID).

Não precisa de fbclient.dll nem de nada instalado: o driver (`firebirdsql`) fala direto com o servidor pela rede. Funciona com Firebird 2.5, 3, 4 e 5.

## Baixar pronto

Baixe o [`ExBird.exe`](ExBird.exe) (botão Download na página do arquivo) e abra com duplo clique. Não precisa instalar nada. Windows pode mostrar um aviso do SmartScreen por ser um programa novo sem assinatura digital: clique em "Mais informações" e "Executar assim mesmo".

SHA-256: `b7e519b9a8b86ab73784528ef4f646a971eddc5d94835df42153740d6ee192e0`

## Gerar o .exe você mesmo (num PC com Python)

Duplo clique em `gerar_exe.bat`. Sai o `ExBird.exe`, arquivo único. Depois é só copiar o .exe pra qualquer PC e abrir.

Sem gerar o .exe: `abrir.bat` (precisa de `pip install -r requirements.txt` antes).

## Fluxo

1. Escolha o .xlsx e a planilha (cabeçalho na linha 1).
2. Preencha host, porta, banco, usuário e senha e clique em "Conectar e listar tabelas". `ITENS` já vem selecionada, mas dá pra trocar.
3. Escolha a coluna de referência no Excel e no banco (se existir `id` nos dois, já vem marcada).
4. Em "Colunas a atualizar", escolha Excel → Banco. "+ Adicionar coluna" pra várias de uma vez.
5. Rode com "Simular" marcado: executa e dá rollback, mostrando o que seria alterado.
6. Desmarque "Simular" e execute pra gravar (commit único no final). Faça backup do .fdb antes.

## Comportamento

- Banco: caminho do .fdb **no servidor**. Se o banco está no próprio PC, host `localhost`.
- Célula vazia: a linha é pulada. Marque "Célula vazia grava NULL" se quiser o contrário.
- Zeros à esquerda (NCM, EAN): célula de texto vai como está; número com formato `00000000` é reconstruído pelo formato; se o Excel guardou só o número, preencha "Zeros" com o tamanho. A coluna do banco tem que ser VARCHAR/CHAR.
- Números tipo `1.234,56` e datas `dd/mm/aaaa` são convertidos.
- O programa lembra banco, tabela, usuário e planilha (`config.json` ao lado do .exe). A senha não é salva.

## Licença

MIT. Veja o arquivo `LICENSE`.
