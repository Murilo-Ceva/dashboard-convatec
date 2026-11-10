# Regras deste projeto

- PowerShell: nunca usar `-EncodedCommand` ou `-enc`. O Defender da CEVA marca como suspeito. Use argumento normal, arquivo `.ps1` ou Python.
- Evitar janela oculta (`-WindowStyle Hidden`) quando houver alternativa simples.
- Scripts `.ps1` só com caracteres ASCII: o PowerShell 5.1 lê UTF-8 sem BOM como ANSI.
- Python com `-I`. Se precisar de biblioteca do usuário: `sys.path.append(site.getusersitepackages())`.
- Sem segredo em arquivo: nada de senha, token ou chave no código, no HTML ou em CSV.
- Dados reais ficam em `dados/`, que o Git ignora. No Git só vai `exemplo.csv`, fictício.
- Repositório Git fora do OneDrive. Pastas `.git` dentro do OneDrive já apareceram vazias.
- Não enviar e-mail, publicar nem fazer push sem pedido explícito.
- Antes de entregar: rodar `scripts/verificar.py` e bater o total com a fonte.
- Dizer o que é oficial, o que é simulado e o que é inferido. Lacuna fica visível no resultado.
