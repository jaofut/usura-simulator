Esse sistema registra empréstimos (quem emprestou, quem recebeu, valor e taxa mensal),
persiste tudo no **Google Sheets** na nuvem (com suporte a cache e backup local) e controla
o saldo devedor com juros simples diários, derivados da própria taxa mensal informada.

### Configuração do Google Sheets (`v01.py`):
1. Crie ou utilize uma Conta de Serviço (Service Account) no Google Cloud Console com acesso à Google Sheets API.
2. Salve o arquivo JSON de chaves baixado como **`credentials.json`** na pasta `usura simulator/` OU configure-o no arquivo `.streamlit/secrets.toml`.
3. Crie uma planilha no Google Drive com o nome **`Gerenciador de Empréstimos - Usura Simulator`** e compartilhe-a dando permissão de **Editor** para o e-mail (`client_email`) da Service Account.
*(Enquanto a conexão não for configurada, o sistema executará normalmente utilizando a memória temporária da sessão e permitindo exportar backups em CSV/XLSX).*

Modelo de juros (importante para entender os cálculos abaixo):
- Juros simples diários incidem sobre o SALDO DE CAPITAL em aberto, à
  taxa diária = taxa_mensal / 100 / 30.
- Isso é calculado de forma contínua desde a data do empréstimo, então em
  30 dias (1 mês) sem nenhum pagamento, o total acumulado é exatamente
  capital x taxa_mensal — ou seja, o "montante" (capital + juros do mês)
  cai certinho no vencimento (empréstimo + 1 mês).
- Cada pagamento é registrado como um evento que "fecha a conta" até
  aquela data: primeiro abate os juros que já tinham acumulado e ainda
  não foram pagos: (Juros_Travados), e só o que sobrar do pagamento abate
  o saldo de capital (Saldo_Principal). A partir daí, os juros voltam a
  contar do zero, sobre o novo saldo de capital.
  Isso é o que corrige o bug: um pagamento de R$ 1000 sobre uma dívida de
  R$ 1500 (capital + juros) deixa exatamente R$ 500 de saldo devedor, e
  não zera o empréstimo.

Status possíveis:
- "Em dia": ainda não passou 1 mês desde o empréstimo, nada pago ainda.
- "Em atraso": passou o vencimento e nada foi pago ainda.
- "Parcialmente pago": já recebeu algum pagamento, mas ainda há saldo.
- "Pago": saldo de capital chegou a zero.
