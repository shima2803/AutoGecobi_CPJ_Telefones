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


def load_credentials(path=CRED_PATH):
    ns = {}
    base_globals = {}
    try:
        import pymysql
        base_globals["pymysql"] = pymysql
    except ImportError:
        class Dummy:
            class cursors:
                class Cursor:
                    pass
        base_globals["pymysql"] = Dummy
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


def conectar_bd_telefones(path=CRED_PATH):
    creds = load_credentials(path)
    cfg = creds.get("BD_TELEFONES_SQLAUTH") or creds.get("BD_TELEFONES_WINDOWS")
    if not cfg:
        raise ValueError("BD_TELEFONES_SQLAUTH ou BD_TELEFONES_WINDOWS não definidos no TXT.")
    driver_name = cfg.get("driver")
    server = cfg.get("server")
    database = cfg.get("database")
    user = cfg.get("user") or creds.get("BD_TELEFONES_USER")
    password = cfg.get("password") or creds.get("BD_TELEFONES_PASS")
    if not user or not password:
        raise ValueError("Usuário/senha do BD_TELEFONES não definidos no TXT.")
    conn_str = (
        f"DRIVER={{{driver_name}}};"
        f"SERVER={server};"
        f"DATABASE={database};"
        f"UID={user};"
        f"PWD={password};"
        "TrustServerCertificate=yes;"
    )
    print(f"Conectando ao BD_TELEFONES: {server} | DB={database} | USER={user}")
    return pyodbc.connect(conn_str)


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


def abrir_gecobi_com_cpj():
    usuario, senha = load_cpj_credentials()
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get(URL_LOGIN)
    wait = WebDriverWait(driver, 30)

    u = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[aria-label='Usuário']")))
    time.sleep(0.5)
    u.clear()
    u.send_keys(usuario)
    u.send_keys(Keys.TAB)

    p = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[aria-label='Senha']")))
    time.sleep(0.5)
    p.clear()
    p.send_keys(senha)
    p.send_keys(Keys.TAB)
    time.sleep(0.5)
    p.send_keys(Keys.ENTER)

    c = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "//i[contains(@class,'material-icons') and normalize-space()='calculate']",
            )
        )
    )
    time.sleep(0.5)
    c.click()

    o = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//div[contains(@class,'q-item__label') and normalize-space()='Operação']")
        )
    )
    time.sleep(0.5)
    o.click()

    t = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "//div[contains(@class,'q-item__section') and contains(@class,'q-item__section--main') and normalize-space()='Telefones']",
            )
        )
    )
    time.sleep(0.5)
    t.click()

    time.sleep(2)
    driver.switch_to.default_content()
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe.frame-legado")))

    c1 = find_checkbox(driver, "l[desbloqueio]", "/html/body/table[2]/tbody/tr[4]/td[2]/font[1]/input")
    if c1:
        driver.execute_script("arguments[0].scrollIntoView(true);", c1)
        time.sleep(0.2)
        driver.execute_script("arguments[0].click();", c1)

    c2 = find_checkbox(driver, "l[lib_processamento]", "/html/body/table[2]/tbody/tr[5]/td[2]/font[1]/input")
    if c2:
        driver.execute_script("arguments[0].scrollIntoView(true);", c2)
        time.sleep(0.2)
        driver.execute_script("arguments[0].click();", c2)

    return driver


driver = abrir_gecobi_com_cpj()
conn = conectar_bd_telefones()
cursor = conn.cursor()

cursor.execute(
    """
    SELECT
        CONCAT(TE.ddd, TE.Numero) AS Telefone,
        p.CpfCnpj,
        CASE
            WHEN TE.ScoreFinal > 80 THEN 'HOT'
            WHEN TE.ScoreFinal >= 60 THEN 'ALTA'
            WHEN TE.ScoreFinal >= 40 THEN 'MEDIA'
            WHEN TE.ScoreFinal < 20 AND TE.Discado = 1 THEN 'IMPRODUTIVO'
            ELSE 'PEQUENA'
        END AS Classificacao,
        TE.ScoreFinal
    FROM dbo.Telefone TE
    JOIN PessoaTelefone PT ON PT.IdTelefone = TE.IdTelefone
    JOIN Pessoa p ON p.IdPessoa = PT.IdPessoa;
    """
)

rows = cursor.fetchall()

class_map = {}
for tel, cpfcnpj, classif, score in rows:
    key = str(classif).strip().upper()
    tel_str = str(tel).strip()
    cpf_str = str(cpfcnpj).strip()
    if tel_str and cpf_str:
        class_map.setdefault(key, []).append(f"{tel_str};{cpf_str}")

status_config = [
    ("IMPRODUTIVO", "3"),
    ("HOT", "2"),
    ("ALTA", "4"),
    ("MEDIA", "5"),
    ("PEQUENA", "6"),
]

for class_name, select_value in status_config:
    lista = class_map.get(class_name, [])
    print(f"Status {class_name}: {len(lista)} registros")
    if not lista:
        continue

    texto_fones = "\n".join(lista)

    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(
        EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe.frame-legado"))
    )

    combo = find_element_any_frame(driver, By.NAME, "l[stnew]", max_depth=6)
    if combo:
        driver.execute_script("arguments[0].scrollIntoView(true);", combo)
        time.sleep(0.2)
        Select(combo).select_by_value(select_value)
        print(f"Status '{class_name}' selecionado no combo.")

    textarea = find_element_any_frame(driver, By.NAME, "l[fones]", max_depth=6)
    if not textarea:
        textarea = find_element_any_frame(
            driver, By.XPATH, "/html/body/table[2]/tbody/tr[6]/td[2]/textarea", max_depth=6
        )

    if textarea and texto_fones:
        driver.execute_script("arguments[0].scrollIntoView(true);", textarea)
        driver.execute_script("arguments[0].value = arguments[1];", textarea, texto_fones)
        driver.execute_script("arguments[0].dispatchEvent(new Event('input'));", textarea)
        print(f"Telefones/CPFs colados para status {class_name}.")

    chk_executar = find_element_any_frame(driver, By.NAME, "l[executar]", max_depth=6)
    if chk_executar:
        driver.execute_script("arguments[0].scrollIntoView(true);", chk_executar)
        time.sleep(0.1)
        driver.execute_script("arguments[0].click();", chk_executar)
        print("Checkbox 'executar' marcado.")
    else:
        print("Checkbox 'executar' não encontrado.")

    btn_executar = find_element_any_frame(driver, By.ID, "confirmaForm", max_depth=6)
    if btn_executar:
        driver.execute_script("arguments[0].scrollIntoView(true);", btn_executar)
        time.sleep(0.1)
        driver.execute_script("arguments[0].click();", btn_executar)
        print(f"Botão 'Executar >>' clicado para status {class_name}.")
    else:
        print("Botão 'Executar >>' não encontrado.")

    time.sleep(2)

print("Processo concluído para todos os status.")
time.sleep(1)

try:
    conn.close()
    print("Conexão com o banco encerrada.")
except:
    pass

try:
    driver.quit()
    print("Navegador fechado.")
except:
    pass

print("Processo finalizado.")
time.sleep(1)
