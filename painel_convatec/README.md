# Painel de follow-up ConvaTec

Copiado do painel HPE v2 (Desktop\HPE\BASE OFICIAL HPE\painel_hpe_v2). Mudam o logo (painel/logo_convatec*.png), os textos (Dealer virou Destinatário) e a base: dados/CONVATEC.csv, exportada da aba Base por NF Oficial da planilha Base CONVATEC - com tracking.xlsm. O template compara DE x PARA; aqui o tema é a situação de entrega por NF, então ficaram o estilo, os filtros que se cruzam, a leitura do recorte, a matriz, o detalhe e a verificação, e o resto foi adaptado.

## Atualizar

1. Coloque a base em `dados/` (CSV ou CSV.gz, com as colunas da aba "Base por NF Oficial") e ajuste `entrada` em `config.json`.
2. Na raiz desta pasta:

```bash
python -I scripts/gerar_painel.py
```

```bash
python -I scripts/verificar.py
```

3. Abra `saida/painel_interno.html` (equipe CEVA) ou `saida/painel_cliente.html` (ConvaTec). Funcionam com duplo clique, sem servidor e sem internet.

Sem Python, o botão **Carregar base** abre outro CSV ou CSV.gz direto no navegador.

## Regras

- Situação: previsão original (PREVISAO DE ENTREGA) contra a data da entrega. Sem entrega, previsão contra a data da base (data do arquivo), que aparece no topo.
- OTD só conta notas entregues. Meta em `config.json` (`meta_otd`).
- Visão cliente: transportadora não entra no arquivo. O frete aparece desde 09/11/2026, a pedido da operação. Não é senha: os dados simplesmente não estão lá.
- Nenhum e-mail entra no HTML (a coluna resp cobrança fica de fora).
