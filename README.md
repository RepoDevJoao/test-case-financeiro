# Case Técnico — Modelagem Financeira
**Gestão de Passivos e Projeção de Fluxo de Caixa (CDI)**

---

## Resumo Executivo

Este projeto simula o fluxo de caixa de uma dívida indexada ao CDI com capitalização
composta em dias úteis (Base 252), considerando pagamentos mensais com prioridade
para juros acumulados.

O modelo foi desenvolvido em Python e contempla:
- Capitalização diária apenas em dias úteis
- Calendário com feriados nacionais brasileiros
- Regra de amortização com priorização de juros
- Geração automatizada de relatório em Excel formatado

---

## Estrutura do Projeto

```
case_financeiro/
├── script.py             # Simulação principal
├── relatorio_mensal.xlsx # Relatório gerado automaticamente
├── requirements.txt      # Dependências Python
└── README.md             # Este arquivo
```

---

## Como Executar

### 1. Instalar dependências
```bash
pip install -r requirements.txt
```

### 2. Rodar a simulação
```bash
python script.py
```

O script imprime o relatório no terminal e gera o arquivo `relatorio_mensal.xlsx` na mesma pasta.

---

## Premissas do Contrato

| Parâmetro           | Valor                      |
|---------------------|----------------------------|
| Principal           | R$ 3.000.000,00            |
| Data de Liberação   | 01/10/2025                 |
| Taxa CDI            | 0,05% ao dia útil          |
| Taxa Fixa           | 0,02% ao dia útil          |
| Parcela Mensal      | R$ 150.000,00              |
| Data de Pagamento   | 1º dia útil de cada mês    |
| Prazo               | 12 meses (nov/2025–out/2026) |

---

## Metodologia

### Calendário — Base 252
O padrão do mercado financeiro brasileiro considera exclusivamente **dias úteis**:
sem sábados, domingos ou feriados nacionais. O calendário de feriados é obtido
via biblioteca `holidays` (feriados nacionais de 2025 e 2026).

A capitalização começa no **dia útil seguinte** à data de liberação (02/10/2025),
pois 01/10/2025 é a data-base do saldo inicial. Essa abordagem segue a prática
de mercado, onde a incidência de juros ocorre apenas após a disponibilização
efetiva do recurso.

### Task 1 — Capitalização Diária
A cada dia útil, o Saldo Devedor é atualizado pelo fator composto:

```
ft = (1 + i_CDI) × (1 + s_fixa)
   = (1 + 0,0005) × (1 + 0,0002)
   = 1,0007001
```

O juro do dia é a diferença nominal: `Jt = SDt − SDt−1`

### Task 2 — Regra de Amortização
No 1º dia útil de cada mês, a parcela de R$ 150.000,00 segue esta hierarquia:

1. **Quita o Juro Acumulado** — soma de todos os `Jt` desde o último pagamento, **inclusive o dia do pagamento**
2. **Amortiza o Principal** — `Amortização = Parcela − J_acum`
   - Se `J_acum > Parcela`: amortização é negativa e o saldo devedor **cresce**
   - Se `Saldo ≤ Parcela` no mês final: a dívida é encerrada pelo valor exato do saldo

---

## Resultados (12 meses)

| Métrica               | Valor              |
|-----------------------|--------------------|
| Total de Juros Pagos  | R$ 471.047,13      |
| Total Amortizado      | R$ 1.328.952,87    |
| Total Pago (parcelas) | R$ 1.800.000,00    |
| Saldo Devedor Final   | R$ 2.142.094,24    |

> A dívida não é integralmente quitada em 12 meses com a parcela de R$ 150.000,00,
> pois o juro mensal médio (Entre R$ 39.000 e R$ 49.000) deixa uma amortização efetiva
> de apenas ~R$ 100.000–R$ 117.000 por período.

---

## Dependências

| Biblioteca | Versão  | Finalidade                                        |
|------------|---------|---------------------------------------------------|
| `holidays` | ≥ 0.68  | Feriados nacionais brasileiros (Base 252)         |
| `pandas`   | ≥ 2.0.0 | Estruturação dos dados e exportação para Excel    |
| `openpyxl` | ≥ 3.1.0 | Formatação do relatório `.xlsx`                   |
