import argparse
import os
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from getpass import getpass
from pathlib import Path
from typing import Optional

import pandas as pd
import pyautogui as py


# =========================
# Configurações (ajuste rápido)
# =========================

DEFAULT_PORTAL_URL = os.getenv(
    "PORTAL_URL",
    "https://example.com/login"  # coloque a URL real do portal aqui (ou via env PORTAL_URL)
)

# Coordenadas da automação (AJUSTE conforme seu monitor/resolução)
# Dica: use uma ferramenta de pegar coordenadas do mouse ou o pyautogui.displayMousePosition()
@dataclass
class ClickMap:
    # Exemplo: clique para focar no campo/aba de seleção de filial
    select_branch: tuple[int, int] = (679, 335)

    # Exemplo: clique no campo data início / data fim (ou onde o portal pede)
    start_date_field: tuple[int, int] = (682, 265)

    # Exemplo: botões/abas do portal que você clica para gerar/baixar relatório
    menu_1: tuple[int, int] = (40, 292)
    menu_2: tuple[int, int] = (31, 333)
    generate_button: tuple[int, int] = (952, 266)
    download_button: tuple[int, int] = (270, 709)


# Mapeamento de filiais (genérico/sanitizado)
BRANCHES = {
    34: "FILIAL 34",
    43: "FILIAL 43",
    44: "FILIAL 44",
}


# =========================
# Utilidades
# =========================

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="RPA (PyAutoGUI) para baixar Excel de portal web e consolidar valores por data."
    )
    parser.add_argument("--start-date", required=True, help="Data início no formato YYYY-MM-DD")
    parser.add_argument("--end-date", required=True, help="Data fim no formato YYYY-MM-DD")
    parser.add_argument("--branch-code", type=int, required=True, help="Código da filial (ex: 34)")
    parser.add_argument(
        "--downloads-dir",
        default=str(Path.home() / "Downloads"),
        help="Pasta onde o arquivo será baixado (default: ~/Downloads)",
    )
    parser.add_argument(
        "--output",
        default=str(Path("output") / "resumo.csv"),
        help="Arquivo de saída (CSV) com consolidação (default: output/resumo.csv)",
    )
    parser.add_argument(
        "--wait-download-seconds",
        type=int,
        default=20,
        help="Tempo máximo (segundos) para aguardar o download aparecer (default: 20)",
    )
    return parser.parse_args()


def safe_get_credentials() -> tuple[str, str]:
    user = os.getenv("PORTAL_USER") or input("Usuário do portal (ou defina PORTAL_USER): ").strip()
    pwd = os.getenv("PORTAL_PASS") or getpass("Senha do portal (ou defina PORTAL_PASS): ").strip()

    if not user or not pwd:
        raise ValueError("Usuário/senha não informados. Use PORTAL_USER/PORTAL_PASS ou digite no prompt.")
    return user, pwd


def most_recent_excel(downloads_dir: Path, after_ts: float) -> Optional[Path]:
    exts = (".xlsx", ".xls")
    candidates = []
    for p in downloads_dir.iterdir():
        if p.is_file() and p.suffix.lower() in exts:
            try:
                if p.stat().st_mtime >= after_ts:
                    candidates.append(p)
            except OSError:
                continue

    if not candidates:
        return None
    return max(candidates, key=lambda x: x.stat().st_mtime)


# =========================
# RPA (PyAutoGUI)
# =========================

def run_rpa(
    portal_url: str,
    username: str,
    password: str,
    branch_name: str,
    start_date: str,
    end_date: str,
    clicks: ClickMap,
) -> None:
    py.PAUSE = 1

    # Abre o Chrome
    py.press("win")
    time.sleep(1)
    py.write("chrome")
    py.press("enter")
    time.sleep(4)

    # Acessa o portal
    py.write(portal_url)
    py.press("enter")
    time.sleep(4)

    # Login (assumindo: campo usuário -> TAB -> campo senha -> TAB -> ENTER)
    py.write(username)
    py.press("tab")
    py.write(password)
    py.press("tab")
    py.press("enter")
    time.sleep(4)

    # Seleciona filial (fluxo baseado no seu script original)
    py.click(*clicks.select_branch)
    py.press("tab")
    py.press("tab")
    py.write(branch_name)

    # Datas
    py.click(*clicks.start_date_field)
    py.hotkey("ctrl", "a")
    py.press("del")
    py.write(start_date)

    py.press("tab")
    py.hotkey("ctrl", "a")
    py.press("del")
    py.write(end_date)

    # Ações no portal para gerar e baixar
    py.click(*clicks.menu_1)
    py.click(*clicks.menu_2)
    py.click(*clicks.generate_button)
    time.sleep(1)
    py.click(*clicks.download_button)


# =========================
# Tratamento do Excel
# =========================

def load_and_summarize(excel_path: Path) -> pd.DataFrame:
    df = pd.read_excel(excel_path)

    # Limpeza básica semelhante ao notebook
    df = df.dropna()

    # Se a primeira linha for header “real” (como no seu notebook), ajusta:
    df.columns = df.iloc[0]
    df = df[1:].copy()

    columns_to_drop = [
        "Descrição", "Moeda", "Couried ID", "Register ID", "Register Name",
        "Bar Code", "Strap Seal Code", "Courier Name",
    ]
    # Remove apenas as que existirem (pra não quebrar caso o layout mude)
    df = df.drop([c for c in columns_to_drop if c in df.columns], axis=1, errors="ignore")

    if "Data" not in df.columns or "Quantia" not in df.columns:
        raise ValueError("Colunas esperadas não encontradas: precisa existir 'Data' e 'Quantia'.")

    df["Data"] = pd.to_datetime(df["Data"], errors="coerce").dt.date
    df = df.dropna(subset=["Data"])

    # Quantia pode vir como texto; tenta converter
    df["Quantia"] = pd.to_numeric(df["Quantia"], errors="coerce")
    df = df.dropna(subset=["Quantia"])

    resumo = df.groupby("Data")["Quantia"].sum().reset_index()
    resumo = resumo.sort_values("Data")
    return resumo


def main() -> int:
    args = parse_args()

    # Validação simples de filial
    if args.branch_code not in BRANCHES:
        print(f"Código de filial não encontrado. Disponíveis: {sorted(BRANCHES.keys())}")
        return 1

    # Credenciais e diretórios
    username, password = safe_get_credentials()
    downloads_dir = Path(args.downloads_dir).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Marca o momento pra capturar o “arquivo mais recente” depois do download
    start_ts = time.time()

    # Executa RPA
    clicks = ClickMap()
    run_rpa(
        portal_url=DEFAULT_PORTAL_URL,
        username=username,
        password=password,
        branch_name=BRANCHES[args.branch_code],
        start_date=args.start_date,
        end_date=args.end_date,
        clicks=clicks,
    )

    # Aguarda download aparecer
    excel_path = None
    deadline = time.time() + args.wait_download_seconds
    while time.time() < deadline:
        excel_path = most_recent_excel(downloads_dir, after_ts=start_ts)
        if excel_path:
            break
        time.sleep(1)

    if not excel_path:
        print("Não encontrei um arquivo Excel baixado recentemente na pasta de downloads.")
        print(f"Pasta verificada: {downloads_dir}")
        return 2

    # Processa e salva
    resumo = load_and_summarize(excel_path)
    resumo.to_csv(output_path, index=False, encoding="utf-8-sig")

    print("\nResumo por data (consolidado):")
    print(resumo.to_string(index=False))
    print(f"\nArquivo salvo em: {output_path}")
    print(f"Excel de origem (detectado): {excel_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
