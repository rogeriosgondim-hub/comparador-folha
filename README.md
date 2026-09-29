# Comparador de Folha de Pagamento

Aplicação web estática para comparar um **Excel original** com um **PDF de folha de pagamento** diretamente no navegador.

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
