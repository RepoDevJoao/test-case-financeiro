"""
Case Técnico: Modelagem Financeira
Gestão de Passivos e Projeção de Fluxo de Caixa (CDI)
"""

# IMPORTAÇÕES
import holidays
import pandas as pd
from datetime import date, timedelta
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# 1. PARÂMETROS DO CONTRATO
PRINCIPAL    = 3_000_000.00   # Aporte inicial (R$)
DATA_INICIO  = date(2025, 10, 1)
I_CDI        = 0.0005         # 0,05% ao dia útil
S_FIXA       = 0.0002         # 0,02% ao dia útil
FATOR_DIARIO = (1 + I_CDI) * (1 + S_FIXA)
PARCELA      = 150_000.00     # Parcela mensal fixa (R$)
MESES        = 12


# 2. CALENDÁRIO (Base 252 — padrão mercado financeiro brasileiro)
def eh_dia_util(d: date) -> bool:
    """Dia útil = não é fim de semana e não é feriado nacional brasileiro."""
    return d.weekday() < 5 and d not in holidays.Brazil(years=d.year)

# Função para encontrar o 1º dia útil de um mês, avançando se necessário
def primeiro_dia_util_do_mes(ano: int, mes: int) -> date:
    """Retorna o 1º dia útil de um mês, avançando se necessário."""
    d = date(ano, mes, 1)
    while not eh_dia_util(d):
        d += timedelta(days=1)
    return d

# Gerar as datas de pagamento: 1º dia útil de cada mês subsequente ao início
def gerar_datas_pagamento(data_inicio: date, n_meses: int) -> list[date]:
    """Gera as n datas de pagamento (1º dia útil de cada mês subsequente)."""
    datas, mes, ano = [], data_inicio.month + 1, data_inicio.year
    for _ in range(n_meses):
        if mes > 12:
            mes, ano = 1, ano + 1
        datas.append(primeiro_dia_util_do_mes(ano, mes))
        mes += 1
    return datas


# 3. SIMULAÇÃO DO FLUXO DE CAIXA DA DÍVIDA
def simular() -> list[dict]:
    """
    Simulação diária do fluxo de caixa da dívida.

    Task 1 — A cada dia útil:
        SDt = SDt-1 * ft,  onde ft = (1 + i_CDI) * (1 + s_fixa)
        Jt  = SDt - SDt-1

    Task 2 — No 1º dia útil de cada mês:
        1. J_acum = soma de todos os Jt desde o último pagamento (inclusive)
        2. A parcela quita J_acum primeiro
        3. Amortização = parcela - J_acum (pode ser negativa se juros > parcela,
           nesse caso o saldo devedor cresce, conforme enunciado)
    """
    datas_pagamento = gerar_datas_pagamento(DATA_INICIO, MESES)
    set_pagamentos  = set(datas_pagamento)

    # Variáveis
    saldo_devedor    = PRINCIPAL
    juros_acumulados = 0.0
    relatorio        = []

    # Capitalização começa no dia seguinte à liberação
    data_atual = DATA_INICIO + timedelta(days=1)

    while data_atual <= datas_pagamento[-1]:
        # Se não for dia útil, pula para o próximo dia
        if not eh_dia_util(data_atual):
            data_atual += timedelta(days=1)
            continue

        # Task 1: capitaliza o saldo devedor e acumula os juros do dia
        sd_anterior      = saldo_devedor
        saldo_devedor    = saldo_devedor * FATOR_DIARIO
        juros_acumulados += saldo_devedor - sd_anterior

        # Task 2: se for dia de pagamento, o 1° dia útil do mês, processa a parcela
        if data_atual in set_pagamentos:
            if saldo_devedor <= PARCELA:
                # Último pagamento: se o saldo devedor for menor que a parcela, paga o que resta e quita a dívida
                amortizacao   = saldo_devedor - juros_acumulados
                parcela_paga  = saldo_devedor
                saldo_devedor = 0.0
            # Se os juros acumulados forem maiores que a parcela, o saldo devedor aumenta (juros não pagos)
            else:
                amortizacao   = PARCELA - juros_acumulados   # pode ser negativa
                parcela_paga  = PARCELA
                saldo_devedor -= amortizacao

            relatorio.append({
                "Data do Pagamento":            data_atual,
                "Juros Acumulados (R$)":        round(juros_acumulados, 2),
                "Amortização do Principal (R$)": round(amortizacao, 2),
                "Parcela (R$)":                 round(parcela_paga, 2),
                "Saldo Devedor (R$)":           round(saldo_devedor, 2),
            }) # Arredonda para 2 casas decimais, como é padrão financeiro

            juros_acumulados = 0.0

            if saldo_devedor <= 0:
                break

        data_atual += timedelta(days=1)

    return relatorio


# 4. EXPORTAÇÃO PARA XLSX COM FORMATAÇÃO
def exportar_xlsx(relatorio: list[dict], caminho: str = "relatorio_mensal.xlsx"):
    df = pd.DataFrame(relatorio)

    totais = {
        "Data do Pagamento":            "TOTAL",
        "Juros Acumulados (R$)":        df["Juros Acumulados (R$)"].sum(),
        "Amortização do Principal (R$)": df["Amortização do Principal (R$)"].sum(),
        "Parcela (R$)":                 df["Parcela (R$)"].sum(),
        "Saldo Devedor (R$)":           "",
    }
    df = pd.concat([df, pd.DataFrame([totais])], ignore_index=True)

    with pd.ExcelWriter(caminho, engine="openpyxl", datetime_format="DD/MM/YYYY") as writer:
        df.to_excel(writer, index=False, sheet_name="Relatório Mensal")

    wb = load_workbook(caminho)
    ws = wb["Relatório Mensal"]

    header_fill = PatternFill("solid", fgColor="2E75B6")
    total_fill  = PatternFill("solid", fgColor="D6E4F0")
    zebra_fill  = PatternFill("solid", fgColor="EBF3FB")
    borda = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"),
    )
    fmt_moeda = "R$ #,##0.00"

    larguras = [20, 24, 28, 16, 22]
    for i, larg in enumerate(larguras, start=1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = larg

    n_linhas = ws.max_row

    for row in ws.iter_rows():
        for cell in row:
            cell.border    = borda
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if cell.row == 1:
                cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
                cell.fill = header_fill
            elif cell.row == n_linhas:
                cell.font = Font(name="Arial", bold=True, size=10)
                cell.fill = total_fill
                if isinstance(cell.value, float):
                    cell.number_format = fmt_moeda
            else:
                cell.font = Font(name="Arial", size=10)
                if cell.row % 2 == 0:
                    cell.fill = zebra_fill
                if isinstance(cell.value, float):
                    cell.number_format = fmt_moeda

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    wb.save(caminho)
    print(f"Relatório salvo em: {caminho}")


# 5. EXECUÇÃO DO SCRIPT E GERAÇÃO DO RELATÓRIO
if __name__ == "__main__":
    relatorio = simular()
    df = pd.DataFrame(relatorio)

    print("\n" + "=" * 92)
    print(f"{'RELATÓRIO MENSAL — FLUXO DE CAIXA DA DÍVIDA':^92}")
    print("=" * 92)
    print(df.to_string(index=False, float_format=lambda x: f"{x:,.2f}"))
    print("=" * 92)
    print(f"\nTotal juros pagos:   R$ {df['Juros Acumulados (R$)'].sum():>14,.2f}")
    print(f"Total amortizado:    R$ {df['Amortização do Principal (R$)'].sum():>14,.2f}")
    print(f"Total pago:          R$ {df['Parcela (R$)'].sum():>14,.2f}")
    print(f"Saldo devedor final: R$ {relatorio[-1]['Saldo Devedor (R$)']:>14,.2f}")

    exportar_xlsx(relatorio)
