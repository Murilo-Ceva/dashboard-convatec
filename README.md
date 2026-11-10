# CEVA · ConvaTec

Código da operação ConvaTec. Os dados reais ficam fora deste repositório, no OneDrive em `Desktop\COVATEC`.

- `macros/`: módulos VBA da `Base CONVATEC - com tracking.xlsm`. `modBaseWV_v12_CONVATEC.bas` atualiza a base a partir do 238 e do BI. Desde a v12e ela acumula pelo CT-e. `modTrackingCONVATEC.bas` cria os rascunhos de follow no Outlook e nunca envia. Os e-mails ficam nas abas da planilha, não no código. Desde a v12f, a planilha tem a aba Agendamentos, com os clientes que exigem agendamento. Na Base por NF Oficial, AC (CLIENTE AGENDADO), AD (CONTATO) e AE (AGENDAMENTO REALIZADO) saem dessa aba. O botão 4 cria os rascunhos de agendamento.
- `painel_convatec/`: painel HTML de follow-up, copiado do painel HPE v2. `dados/exemplo.csv` é fictício. Veja o README da pasta.

## Painel no site

`index.html` é o painel **sem nenhuma nota**. Abra e use **Carregar base** com o `CONVATEC.csv` (export da aba Base por NF Oficial). O arquivo é lido só no navegador e não sai do computador. Os destinatários são pacientes (dados de saúde, LGPD), então nenhuma base entra neste repositório público.
