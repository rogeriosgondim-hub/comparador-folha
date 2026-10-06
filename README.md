# Folha de Pagamento — Comparador e Preenchimento

Aplicação web estática com duas funções: comparar um **Excel original** com um **PDF de folha de pagamento** e preencher automaticamente a coluna **Remuneração líquida a receber** no Excel a partir do PDF.

## Privacidade

- Os arquivos selecionados não são enviados para servidor.
- O processamento acontece em JavaScript na própria aba do navegador.
- O projeto não possui backend, banco de dados ou API de upload.
- SheetJS e PDF.js são carregados por CDN.

## Funções

- Upload de .xlsx, .xls ou .csv.
- Upload do PDF da folha.
- CPF como chave principal.
- Nome normalizado como chave secundária.
- Tolerância configurável, padrão de R$ 0,01.
- Diferenças de centavos separadas.
- Tratamento de TRCT/rescisões.
- Identificação de novos colaboradores, ausentes, duplicados e casos para verificar.
- Painel de resumo.
- Filtros de resultado.
- Exportação das divergências para Excel.
- Módulo de preenchimento automático do líquido no .xlsx.
- CPF como chave principal para o preenchimento; nome normalizado como alternativa.
- Regra específica para TRCT/rescisões, priorizando a rubrica de líquido da rescisão.
- Tela de conferência antes de gerar o arquivo final.
- Exportação da trilha de conferência.
- Preservação do pacote .xlsx: o módulo altera apenas as células de líquido identificadas.

## Estrutura

```text
comparador-folha/
├── .nojekyll
├── index.html
├── README.md
└── assets/
    ├── css/
    │   └── styles.css
    └── js/
        └── app.js
```

## GitHub Pages

No repositório, abra **Settings → Pages**.

Em **Build and deployment**:
1. Source: **Deploy from a branch**
2. Branch: **main**
3. Folder: **/ (root)**
4. Clique em **Save**

O endereço esperado é:

```text
https://rogeriosgondim-hub.github.io/comparador-folha/
```

## Observação sobre PDFs

PDF não é uma tabela estruturada. O parser foi preparado para o padrão de extrato mensal usado no projeto, procurando empregado, CPF, vínculo, líquido e indicações de rescisão/TRCT.

Antes de considerar o resultado definitivo, confira o bloco **Diagnóstico da importação**, especialmente a quantidade de registros identificados no Excel e no PDF.

## Excel com líquidos vazios

Se “Remuneração líquida a receber” estiver vazia, a comparação monetária fica pendente. Use “Preencher líquidos com estes arquivos” para reaproveitar o Excel e o PDF, conferir a prévia e gerar uma nova cópia. Comparar essa cópia com o mesmo PDF valida a transferência dos valores, mas não constitui conferência independente do cálculo da folha.

A competência é extraída exclusivamente do campo “Competência” do PDF; datas de emissão e admissão não são utilizadas.

Teste de regressão com dados fictícios: `node tests/payroll-import.cjs`.
