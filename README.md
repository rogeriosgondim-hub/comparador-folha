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

Os registros encontrados no PDF com líquido vazio no Excel recebem “LÍQUIDO VAZIO NO EXCEL” e informam a célula a preencher. “VERIFICAR” é reservado para dados que não puderam ser reconhecidos. Novos colaboradores e TRCT mantêm suas categorias específicas; o indicador de líquidos vazios inclui todos os registros do Excel.

## Conferência ampliada C a AN

A seção “Conferência por campo — C a AN” analisa as 38 colunas para cada colaborador. A correspondência usa CPF primeiro e nome normalizado como alternativa; duplicidades bloqueiam a associação automática. O cabeçalho de cada coluna é validado antes de aplicar o mapeamento.

- Cadastro: nome, CPF, categoria de vínculo e admissão.
- Salário: campo contratual Salário do PDF versus I (estágio) ou J (funcionário), sem confundir com proventos ou dias normais proporcionais. Aprendiz com remuneração em I e J vazio é tratado explicitamente.
- Rubricas: saúde 339, odonto dependentes 334, home office 416, educação 447, creche 492, empréstimo interno 474; descrições específicas para comissões, bônus, Wellhub e VT. Horas extras/adicional noturno usam referências em horas, sem incluir reflexos DSR.
- Jornada semanal versus mensal, indicadores de benefícios, faltas, observações e crédito consignado/provisões exigem conferência manual.
- Nascimento, idade, custos/subsídios do plano, banco/PIX/agência/conta e plano escolhido não têm equivalente neste extrato.
- Ausência de rubrica não é zero nem aprovação. Célula vazia com valor no PDF é pendência explícita. A soma das rubricas extraídas deve conferir com proventos e descontos de cada bloco para validar comparações de rubricas.

A exportação completa contém cinco abas: Conferência C-AN, Divergências, Pendências, Mapeamento e Rubricas PDF. A tela limita a exibição a blocos de 200 campos, com busca e filtros; a exportação inclui todos. A comparação e o histórico de líquido permanecem identificados separadamente.

Teste específico: `node tests/field-audit.mjs`. Dados de produção nunca entram no repositório.
