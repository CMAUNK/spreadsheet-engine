# SmartSheet v0.1

Aplicação Streamlit para preencher cópias de planilhas `.xlsx` sem modificar o arquivo enviado.

## Executar

```powershell
python -m pip install -r requirements.txt
streamlit run app.py
```

## Nesta versão

- upload de qualquer `.xlsx` e seleção de aba;
- mapa da planilha e detecção do cabeçalho mais provável;
- seleção de colunas por nome, linha inicial e célula opcional de data;
- parser para formatos como `S21 - 6`, `S21: 6` e `Produto A = 15`;
- prévia editável, cópia Excel com formatação preservada e PDF exportado pelo Microsoft Excel;
- salvamento local da configuração para o mesmo arquivo.

O PDF é gerado pelo Microsoft Excel, preservando o layout e as configurações de impressão da aba. Caso esse recurso não esteja disponível, o programa oferece um relatório simples como alternativa.
