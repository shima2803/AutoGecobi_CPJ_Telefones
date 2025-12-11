import os
import time
import pyodbc
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

CRED_PATH = r"\\fs01\ITAPEVA ATIVAS\DADOS\SA_Credencials.txt"
URL_LOGIN = "http://192.168.0.251/gecobi2/qfrontend/#/auth/login"


# =====================================================================
# CARREGAR CREDENCIAIS
# =====================================================================
def load_credentials(path=CRED_PATH):
    ns = {}
    base_globals = {}
    try:
        import pymysql
        base_globals["pymysql"] = pymysql
    except:
        base_globals["pymysql"] = None

    with open(path, "r", encoding="utf-8") as f:
        code = f.read()

    exec(code, base_globals, ns)
    return ns


def load_cpj_credentials(path=CRED_PATH):
    creds = load_credentials(path)
    user = creds.get("CPJ_USER")
    password = creds.get("CPJ_PASS")
    if not user or not password:
        raise ValueError("CPJ_USER ou CPJ_PASS não encontrados no TXT.")
    return user, password


# =====================================================================
# CONECTAR AO BANCO (ATUALIZADO COM AUTO-FALLBACK DE DRIVER)
# =====================================================================
def conectar_bd_telefones(path=CRED_PATH):
    creds = load_credentials(path)

    cfg = creds.get("BD_TELEFONES_SQLAUTH") or creds.get("BD_TELEFONES_WINDOWS")
    if not cfg:
        raise ValueError("BD_TELEFONES_SQLAUTH / BD_TELEFONES_WINDOWS não definidos no TXT.")

    driver_name_cfg = (cfg.get("driver") or "").strip()
    server = (cfg.get("server") or "").strip()
    database = (cfg.get("database") or "").strip()
    user = (cfg.get("user") or creds.get("BD_TELEFONES_USER") or "").strip()
    password = (cfg.get("password") or creds.get("BD_TELEFONES_PASS") or "").strip()

    print("\n== CONFIGURAÇÃO DO BD LIDA NO TXT ==")
    print(cfg)
    print("Driver informado:", repr(driver_name_cfg))

    drivers_disp = [d.strip() for d in pyodbc.drivers()]
    print("Drivers ODBC disponíveis:", drivers_disp)

    # Tenta usar o driver do TXT → senão usa o primeiro "SQL Server" disponível
    if driver_name_cfg in drivers_disp:
        driver_name = driver_name_cfg
    else:
        candidatos = [d for d in drivers_disp if "SQL Server" in d]
        if not candidatos:
            raise RuntimeError(
                f"Nenhum driver SQL Server disponível. Drivers encontrados: {drivers_disp}"
            )
        driver_name = candidatos[0]
        print(f"AVISO: driver '{driver_name_cfg}' não existe nesta máquina. Usando '{driver_name}'.")

    if not server or not database:
        raise RuntimeError(f"Configuração do banco incompleta: server={server}, database={database}")

    if not user or not password:
        raise RuntimeError("Usuário/senha do banco não encontrados no TXT.")

    conn_str = (
        f"DRIVER={{{driver_name}}};"
        f"SERVER={server};"
        f"DATABASE={database};"
        f"UID={user};"
        f"PWD={password};"
        "TrustServerCertificate=yes;"
    )

    print("\nString de conexão (senha oculta):")
    print(conn_str.replace(password, "********"))

    print(f"Conectando ao BD: {server} | DB={database} | USER={user}")
    return pyodbc.connect(conn_str)


# =====================================================================
# FUNÇÕES DE LOCALIZAÇÃO EM FRAMES
# =====================================================================
def find_checkbox(driver, name_value=None, xpath_value=None, max_depth=6, depth=0):
    if depth > max_depth:
        return None

    if name_value:
        for el in driver.find_elements(By.NAME, name_value):
            if el.tag_name.lower() == "input":
                return el

    if xpath_value:
        for el in driver.find_elements(By.XPATH, xpath_value):
            if el.tag_name.lower() == "input":
                return el

    frames = driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe")
    for fr in frames:
        try:
            driver.switch_to.frame(fr)
            el = find_checkbox(driver, name_value, xpath_value, max_depth, depth + 1)
            if el:
                return el
            driver.switch_to.parent_frame()
        except:
            try:
                driver.switch_to.parent_frame()
            except:
                pass

    return None


def find_element_any_frame(driver, by, value, max_depth=6, depth=0):
    if depth > max_depth:
        return None

    try:
        elems = driver.find_elements(by, value)
        if elems:
            return elems[0]
    except:
        pass

    frames = driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe")

    for fr in frames:
        try:
            driver.switch_to.frame(fr)
            el = find_element_any_frame(driver, by, value, max_depth, depth + 1)
            if el:
                return el
            driver.switch_to.parent_frame()
        except:
            try:
                driver.switch_to.parent_frame()
            except:
                pass

    return None


# =====================================================================
# ABRIR Gecobi COM LOGIN + MENU
# =====================================================================
def abrir_gecobi_com_cpj():
    usuario, senha = load_cpj_credentials()
    driver = webdriver.Chrome()
    driver.maximize_window()
    wait = WebDriverWait(driver, 30)

    driver.get(URL_LOGIN)

    # -------- Login --------
    u = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[aria-label='Usuário']")))
    u.send_keys(usuario)
    u.send_keys(Keys.TAB)

    p = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[aria-label='Senha']")))
    p.send_keys(senha)
    p.send_keys(Keys.ENTER)

    # -------- Checkbox extra (quando aparece) --------
    try:
        time.sleep(1.5)
        extra_chk = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "form div.q-checkbox.cursor-pointer"))
        )
        driver.execute_script("arguments[0].click();", extra_chk)
        print("Checkbox extra marcado.")

        # clicar novamente no botão Entrar
        btn_entrar = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//span[@class='block' and normalize-space()='Entrar']/ancestor::button")
            )
        )
        driver.execute_script("arguments[0].click();", btn_entrar)
        print("Botão 'Entrar' clicado novamente.")
    except:
        print("Não apareceu checkbox extra. Seguindo...")

    # -------- Menu superior: calculate --------
    calc = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//i[contains(@class,'material-icons') and normalize-space()='calculate']")
    ))
    driver.execute_script("arguments[0].click();", calc)
    print("Ícone 'calculate' clicado.")

    # -------- Operação --------
    op = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//div[contains(@class,'q-item__label') and normalize-space()='Operação']")
    ))
    driver.execute_script("arguments[0].click();", op)
    print("Menu 'Operação' clicado.")

    # -------- Telefones --------
    tel = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//div[contains(@class,'q-item__section--main') and normalize-space()='Telefones']")
    ))
    driver.execute_script("arguments[0].click();", tel)
    print("Menu 'Telefones' clicado.")

    # -------- Entrar no legado --------
    time.sleep(2)
    driver.switch_to.default_content()
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe.frame-legado")))

    # checkbox desbloqueio
    c1 = find_checkbox(driver, "l[desbloqueio]")
    if c1:
        driver.execute_script("arguments[0].click();", c1)

    # checkbox lib_processamento
    c2 = find_checkbox(driver, "l[lib_processamento]")
    if c2:
        driver.execute_script("arguments[0].click();", c2)

    return driver


# =====================================================================
# PROCESSO PRINCIPAL
# =====================================================================

driver = abrir_gecobi_com_cpj()
conn = conectar_bd_telefones()
cursor = conn.cursor()

cursor.execute("""
    SELECT
        CONCAT(TE.ddd, TE.Numero) AS Telefone,
        p.CpfCnpj,
        CASE
            WHEN TE.ScoreFinal >= 80 THEN 'HOT'
            WHEN TE.ScoreFinal >= 60 THEN 'ALTA'
            WHEN TE.ScoreFinal >= 40 THEN 'MEDIA'
            WHEN TE.ScoreFinal < 20 AND TE.Discado = 1 and te.contato = 0 THEN 'IMPRODUTIVO'
            ELSE 'PEQUENA'
        END AS Classificacao,
        TE.ScoreFinal
    FROM dbo.Telefone TE
    JOIN PessoaTelefone PT ON PT.IdTelefone = TE.IdTelefone
    JOIN Pessoa p ON p.IdPessoa = PT.IdPessoa;
""")

rows = cursor.fetchall()

class_map = {}
for tel, cpf, cls, score in rows:
    class_map.setdefault(cls.upper(), []).append(f"{tel};{cpf}")

status_config = [
    ("IMPRODUTIVO", "3"),
    ("HOT", "2"),
    ("ALTA", "4"),
    ("MEDIA", "5"),
    ("PEQUENA", "6"),
]

for class_name, valor_select in status_config:
    lista = class_map.get(class_name, [])
    print(f"\nStatus {class_name}: {len(lista)} registros")

    if not lista:
        continue

    texto = "\n".join(lista)

    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(
        EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe.frame-legado"))
    )

    # selecionar status
    combo = find_element_any_frame(driver, By.NAME, "l[stnew]")
    if combo:
        Select(combo).select_by_value(valor_select)

    # preencher textarea
    textarea = find_element_any_frame(driver, By.NAME, "l[fones]")
    if textarea:
        driver.execute_script("arguments[0].value = arguments[1];", textarea, texto)
        driver.execute_script("arguments[0].dispatchEvent(new Event('input'));", textarea)

    # checkbox executar
    chk = find_element_any_frame(driver, By.NAME, "l[executar]")
    if chk:
        driver.execute_script("arguments[0].click();", chk)

    # botão executar
    botao = find_element_any_frame(driver, By.ID, "confirmaForm")
    if botao:
        driver.execute_script("arguments[0].click();", botao)

    time.sleep(2)

print("\nProcesso concluído.")

try:
    conn.close()
except:
    pass

try:
    driver.quit()
except:
    pass

print("Finalizado.")
