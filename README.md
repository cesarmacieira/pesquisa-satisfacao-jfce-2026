# pesquisa-satisfacao-jfce

Pesquisa de satisfação quantitativa da Justiça Federal no Ceará (servidores, advogados, magistrados e usuários externos), feita em Streamlit.

## Como as respostas são guardadas

O disco do Streamlit Community Cloud é apagado a cada redeploy/reinício. Por isso, o app grava as respostas em uma base do Airtable em vez de um arquivo local — assim os dados sobrevivem a qualquer redeploy. O Airtable foi escolhido por ser gratuito (sem cartão de crédito) e por não hibernar por inatividade, ao contrário de outras opções gratuitas.

Se as credenciais do Airtable **não** estiverem configuradas (por exemplo, rodando o app na sua máquina sem configurar nada), ele volta a gravar no arquivo local `Dados.xlsx`, exatamente como antes. Ou seja, você pode continuar testando localmente sem precisar configurar nada — a configuração abaixo só é necessária para a versão que fica no ar (deploy).

Essa lógica fica isolada em duas funções do `app.py` (`load_data()` e `append_row()`), então trocar o Airtable por um banco de dados de verdade no futuro (ex.: banco interno da JFCE) é só reescrever essas duas funções — o resto do app não muda.

**Limite do plano gratuito:** 1.000 respostas por base. Para essa primeira versão/piloto é suficiente; se estiver perto do limite, dá para exportar e arquivar as respostas antigas, ou migrar para um banco de verdade.

### Passo a passo para configurar o Airtable (uma vez só)

1. **Crie uma conta gratuita** em [airtable.com](https://airtable.com) (não pede cartão de crédito).

2. **Crie uma base em branco** (qualquer nome). Toda base nova entra automaticamente num trial de 14 dias de recursos pagos, que expira sozinho e volta pro plano gratuito — não é cobrança, pode ignorar o aviso de "Trial". Copie o **Base ID** da URL do navegador (a parte que começa com `app...`, logo após `airtable.com/`).

3. **Crie um Personal Access Token** em [airtable.com/create/tokens](https://airtable.com/create/tokens) → "Create new token":
   - Em **Scopes**, adicione: `data.records:read`, `data.records:write`, `schema.bases:read`, `schema.bases:write`
   - Em **Access**, clique "Add a base" e selecione a base criada no passo 2
   - Copie o token gerado (começa com `pat...`) — ele só aparece uma vez

4. **Crie a tabela `respostas`** na base, com um campo (tipo "Single line text") para cada coluna listada em `COLUMNS` no início do `app.py`. Isso pode ser feito automaticamente via API do Airtable a partir do token e do Base ID (foi assim que a tabela deste projeto foi criada).

5. **Configure os "Secrets" no Streamlit Cloud.** No painel do seu app (share.streamlit.io → seu app → Settings → Secrets), cole:

   ```toml
   airtable_token = "patXXXXXXXXXXXXXX.XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
   airtable_base_id = "appXXXXXXXXXXXXXX"
   ```

6. **Redeploy o app.** Assim que os secrets forem salvos, o Streamlit Cloud reinicia o app automaticamente e ele passa a gravar direto na base do Airtable.

Para testar localmente com o Airtable (em vez do fallback local), crie um arquivo `.streamlit/secrets.toml` na pasta do projeto com o mesmo conteúdo acima — esse arquivo não deve ser commitado no Git (já está no `.gitignore`).
