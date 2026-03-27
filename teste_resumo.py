"""
Teste de leitura de notas finais do RCO.

Execução:
    python teste_resumo.py
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from rco.auth import conectar_chrome
from database import entrar_turma
from rco.notas import buscar_notas_finais_rco
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

ESCOLA     = "ALGATE LICKFELD MAUS, C E PROFA-EF M"
TURMA      = "3ª Série - Noite - A"
DISCIPLINA = "FISICA"
TRIMESTRE  = "1º Tri"


def main():
    print("Conectando ao Chrome...")
    browser = conectar_chrome()
    print(f"Conectado: {browser.title}")

    browser.get("https://rco.paas.pr.gov.br/livro")
    WebDriverWait(browser, 15).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div.card"))
    )

    print(f"\nEntrando na turma: {TURMA} | {DISCIPLINA} | {TRIMESTRE}")
    ok = entrar_turma(browser, ESCOLA, TURMA, DISCIPLINA, TRIMESTRE)
    if not ok:
        print("ERRO: não foi possível entrar na turma.")
        return

    print("Buscando notas finais...\n")
    alunos = buscar_notas_finais_rco(browser)

    print(f"{'Nº':<4} {'Nome':<40} {'Situação':<10} {'Soma'}")
    print("-" * 65)
    for a in alunos:
        soma = str(a["soma"]) if a["soma"] is not None else "-"
        print(f"{a['numero']:<4} {a['nome']:<40} {a['situacao']:<10} {soma}")

    print(f"\nTotal: {len(alunos)} aluno(s)")


if __name__ == "__main__":
    main()
