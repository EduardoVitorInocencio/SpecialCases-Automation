import pyautogui
import pyperclip
import time
import pandas as pd
import os

# =========================================================== Configurações
# Carregar dados do Excel
df = pd.read_excel("dataSources/special_cases.xlsx", sheet_name="data")

# Criar a pasta de logs se não existir
log_dir = "history"
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, "log.txt")

# Definição de tempos para facilitar a alteração
TEMPO_INICIAL = 2  # Tempo antes de iniciar a automação
TEMPO_MOVIMENTO = 0.25  # Tempo de movimentação do mouse
TEMPO_SELECIONAR_TEXTO = 0.2  # Tempo para selecionar o texto no campo
TEMPO_LIMPAR_CAMPO = 0.2  # Tempo após apagar o campo
TEMPO_COPIAR = 1  # Tempo para garantir que o texto foi copiado
TEMPO_COLAR = 0.5  # Tempo após colar o texto
TEMPO_EXTRA_CLIQUE = 2  # Tempo extra para garantir a interação entre campos
TEMPO_PRESS_ESC = 1

# Aguardar um tempo antes de começar
time.sleep(TEMPO_INICIAL)

# Iniciar o timer
tempo_inicio = time.time()
total_contratos = len(df)

# ===========================================================
def preencher_campo(x, y, texto):
    """
    Move o mouse até as coordenadas (x, y), clica, limpa o campo e insere o texto via área de transferência.
    """
    pyautogui.moveTo(x, y, duration=TEMPO_MOVIMENTO)
    pyautogui.click()
    
    # Selecionar e limpar o campo
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(TEMPO_SELECIONAR_TEXTO)
    pyautogui.press('backspace')
    time.sleep(TEMPO_LIMPAR_CAMPO)

    # Copiar e colar o texto
    pyperclip.copy(str(texto))
    time.sleep(TEMPO_COPIAR)
    pyautogui.hotkey('ctrl', 'v')

    time.sleep(TEMPO_COLAR)

# =========================================================== Processamento de contratos
for index, row in df.iterrows():
    contrato_atual = index + 1
    restantes = total_contratos - contrato_atual

    pyautogui.moveTo(1700, 409, duration=TEMPO_MOVIMENTO)
    pyautogui.click()

    time.sleep(TEMPO_PRESS_ESC)
    pyautogui.hotkey('esc')

    time.sleep(TEMPO_PRESS_ESC)
    pyautogui.hotkey('esc')

    time.sleep(TEMPO_PRESS_ESC)
    pyautogui.hotkey('esc')

    time.sleep(TEMPO_PRESS_ESC)
    pyautogui.hotkey('esc')

    time.sleep(TEMPO_PRESS_ESC)
    pyautogui.hotkey('esc')

    # Selecionando Favoritos
    pyautogui.moveTo(1366, 169, duration=TEMPO_MOVIMENTO)
    time.sleep(TEMPO_EXTRA_CLIQUE)
    pyautogui.click()

    # Selecionando Account Display (SFCA)
    pyautogui.moveTo(1366, 207, duration=TEMPO_MOVIMENTO)
    time.sleep(TEMPO_EXTRA_CLIQUE)
    pyautogui.click()

    contract = str(row["contract"])  # Pegando o número do contrato
    new_date = str(row["new_date"])  # Pegando a nova data

    # =========================================================== Preenchendo os campos
    preencher_campo(1606, 206, "333")  # Campo Comp.
    time.sleep(TEMPO_EXTRA_CLIQUE)

    preencher_campo(1604, 241, "113")  # Campo ACC.
    time.sleep(TEMPO_EXTRA_CLIQUE)

    preencher_campo(1656, 238, contract)  # Campo Contract Number.
    time.sleep(TEMPO_EXTRA_CLIQUE)

    # =========================================================== Selecionando o mês
    pyautogui.moveTo(2086, 301, duration=TEMPO_MOVIMENTO)
    time.sleep(TEMPO_EXTRA_CLIQUE)
    pyautogui.click()

    pyautogui.moveTo(2045, 413, duration=TEMPO_MOVIMENTO)
    time.sleep(TEMPO_EXTRA_CLIQUE)
    pyautogui.rightClick()

    pyautogui.hotkey('enter')

    # =========================================================== Clicando na primeira parcela e selecionando Rescheduling
    pyautogui.moveTo(2072, 564, duration=TEMPO_MOVIMENTO)
    time.sleep(TEMPO_EXTRA_CLIQUE)
    pyautogui.rightClick()

    pyautogui.moveTo(2162, 490, duration=TEMPO_MOVIMENTO)
    time.sleep(TEMPO_EXTRA_CLIQUE)
    pyautogui.click()

    # =========================================================== Alterando a Rec Date
    time.sleep(2)
    preencher_campo(2329, 563, new_date)  # Data de Rec Date
    time.sleep(TEMPO_EXTRA_CLIQUE)
    pyautogui.hotkey('enter')
    time.sleep(2)
    pyautogui.hotkey('esc')

    # =========================================================== Registrando no log
    with open(log_file, "a", encoding="utf-8") as log:
        log.write(f"Contrato {contract} atualizado para nova data {new_date}\n")

    # =========================================================== Exibição no terminal
    tempo_atual = time.time()
    tempo_decorrido = tempo_atual - tempo_inicio
    tempo_medio_por_contrato = tempo_decorrido / contrato_atual
    tempo_estimado_restante = tempo_medio_por_contrato * restantes

    print(f"✅ Contrato {contract} atualizado para nova data {new_date}")
    print(f"📊 Progresso: {contrato_atual}/{total_contratos} contratos processados.")
    print(f"⏳ Tempo decorrido: {tempo_decorrido:.2f} segundos")
    print(f"⏱️ Tempo médio por contrato: {tempo_medio_por_contrato:.2f} segundos")
    print(f"⌛ Tempo estimado restante: {tempo_estimado_restante:.2f} segundos\n")

# Tempo total de execução
tempo_total = time.time() - tempo_inicio
print(f"🚀 Processo finalizado em {tempo_total:.2f} segundos!")
