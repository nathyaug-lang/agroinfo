"""Decisões de política monetária — FOMC (Fed Funds Rate) e calendários.

A série da Selic-Meta (Copom) NÃO está aqui — é puxada em tempo real do
SGS/BCB pela função `fetch_selic_history` no build.py (série 432, com
cache local em `.selic_history_cache.json`). Isso garante que os valores
permaneçam fiéis à fonte oficial.

Cada lista de FOMC contém (data_iso, taxa_pct, evento, observacao).
- Para o FOMC pós-dez/2008 a taxa registrada é o ponto médio do range alvo.
- "evento" é "Reunião regular" ou "Reunião extraordinária".
- FOMC consolidado a partir dos comunicados oficiais do Federal Reserve;
  confira no comunicado original antes de uso operacional.
"""

# (data, taxa_pct, evento, obs)
FOMC = [
    ("2003-01-29", 1.25,  "Reunião regular", "Mantém"),
    ("2003-06-25", 1.00,  "Reunião regular", "Corte -25 bp"),
    ("2004-06-30", 1.25,  "Reunião regular", "Início do ciclo de alta · +25 bp"),
    ("2004-08-10", 1.50,  "Reunião regular", "+25 bp"),
    ("2004-09-21", 1.75,  "Reunião regular", "+25 bp"),
    ("2004-11-10", 2.00,  "Reunião regular", "+25 bp"),
    ("2004-12-14", 2.25,  "Reunião regular", "+25 bp"),
    ("2005-02-02", 2.50,  "Reunião regular", "+25 bp"),
    ("2005-03-22", 2.75,  "Reunião regular", "+25 bp"),
    ("2005-05-03", 3.00,  "Reunião regular", "+25 bp"),
    ("2005-06-30", 3.25,  "Reunião regular", "+25 bp"),
    ("2005-08-09", 3.50,  "Reunião regular", "+25 bp"),
    ("2005-09-20", 3.75,  "Reunião regular", "+25 bp"),
    ("2005-11-01", 4.00,  "Reunião regular", "+25 bp"),
    ("2005-12-13", 4.25,  "Reunião regular", "+25 bp"),
    ("2006-01-31", 4.50,  "Reunião regular", "+25 bp"),
    ("2006-03-28", 4.75,  "Reunião regular", "+25 bp"),
    ("2006-05-10", 5.00,  "Reunião regular", "+25 bp"),
    ("2006-06-29", 5.25,  "Reunião regular", "+25 bp · pico do ciclo"),
    ("2007-09-18", 4.75,  "Reunião regular", "Início do ciclo de cortes · -50 bp"),
    ("2007-10-31", 4.50,  "Reunião regular", "-25 bp"),
    ("2007-12-11", 4.25,  "Reunião regular", "-25 bp"),
    ("2008-01-22", 3.50,  "Reunião extraordinária", "-75 bp · emergência"),
    ("2008-01-30", 3.00,  "Reunião regular", "-50 bp"),
    ("2008-03-18", 2.25,  "Reunião regular", "-75 bp"),
    ("2008-04-30", 2.00,  "Reunião regular", "-25 bp"),
    ("2008-10-08", 1.50,  "Reunião extraordinária", "-50 bp · ação coordenada"),
    ("2008-10-29", 1.00,  "Reunião regular", "-50 bp"),
    ("2008-12-16", 0.125, "Reunião regular", "Banda 0,00–0,25% · ZLB"),
    ("2015-12-16", 0.375, "Reunião regular", "+25 bp · saída do ZLB"),
    ("2016-12-14", 0.625, "Reunião regular", "+25 bp"),
    ("2017-03-15", 0.875, "Reunião regular", "+25 bp"),
    ("2017-06-14", 1.125, "Reunião regular", "+25 bp"),
    ("2017-12-13", 1.375, "Reunião regular", "+25 bp"),
    ("2018-03-21", 1.625, "Reunião regular", "+25 bp"),
    ("2018-06-13", 1.875, "Reunião regular", "+25 bp"),
    ("2018-09-26", 2.125, "Reunião regular", "+25 bp"),
    ("2018-12-19", 2.375, "Reunião regular", "+25 bp · pico do ciclo"),
    ("2019-07-31", 2.125, "Reunião regular", "-25 bp · seguro contra desaceleração"),
    ("2019-09-18", 1.875, "Reunião regular", "-25 bp"),
    ("2019-10-30", 1.625, "Reunião regular", "-25 bp"),
    ("2020-03-03", 1.125, "Reunião extraordinária", "-50 bp · COVID-19"),
    ("2020-03-15", 0.125, "Reunião extraordinária", "-100 bp · retorno ao ZLB"),
    ("2022-03-16", 0.375, "Reunião regular", "+25 bp · início do ciclo de alta"),
    ("2022-05-04", 0.875, "Reunião regular", "+50 bp"),
    ("2022-06-15", 1.625, "Reunião regular", "+75 bp"),
    ("2022-07-27", 2.375, "Reunião regular", "+75 bp"),
    ("2022-09-21", 3.125, "Reunião regular", "+75 bp"),
    ("2022-11-02", 3.875, "Reunião regular", "+75 bp"),
    ("2022-12-14", 4.375, "Reunião regular", "+50 bp"),
    ("2023-02-01", 4.625, "Reunião regular", "+25 bp"),
    ("2023-03-22", 4.875, "Reunião regular", "+25 bp"),
    ("2023-05-03", 5.125, "Reunião regular", "+25 bp"),
    ("2023-07-26", 5.375, "Reunião regular", "+25 bp · pico do ciclo"),
    ("2024-09-18", 4.875, "Reunião regular", "-50 bp · início do ciclo de cortes"),
    ("2024-11-07", 4.625, "Reunião regular", "-25 bp"),
    ("2024-12-18", 4.375, "Reunião regular", "-25 bp"),
    ("2025-09-17", 4.125, "Reunião regular", "-25 bp · retomada do ciclo"),
    ("2025-10-29", 3.875, "Reunião regular", "-25 bp"),
    ("2025-12-10", 3.625, "Reunião regular", "-25 bp"),
]


# Próximas reuniões — calendário oficial publicado pelos respectivos comitês.
# Para reuniões de 2 dias, registra-se o dia da decisão (segundo dia).
#
# A coluna "projeção" reflete o consenso de mercado (CME FedWatch para o Fed,
# pesquisa Focus/BCB para o Copom). Atualize manualmente conforme o novo
# consenso for publicado.
#
# Tupla: (data_iso, evento, projecao_str)
PROXIMAS_FOMC = [
    ("2026-06-17", "Reunião regular · projeções (SEP)", "3,375% (consenso · −25 bp)"),
    ("2026-07-29", "Reunião regular",                   "3,375% (consenso · manter)"),
    ("2026-09-16", "Reunião regular · projeções (SEP)", "3,125% (consenso · −25 bp)"),
    ("2026-10-28", "Reunião regular",                   "3,125% (consenso · manter)"),
    ("2026-12-16", "Reunião regular · projeções (SEP)", "2,875% (consenso · −25 bp)"),
]

PROXIMAS_COPOM = [
    ("2026-06-17", "Reunião regular", "14,25% (Focus · −25 bp)"),
    ("2026-07-29", "Reunião regular", "14,00% (Focus · −25 bp)"),
    ("2026-09-16", "Reunião regular", "13,75% (Focus · −25 bp)"),
    ("2026-10-28", "Reunião regular", "13,50% (Focus · −25 bp)"),
    ("2026-12-09", "Reunião regular", "13,25% (Focus · −25 bp)"),
]
