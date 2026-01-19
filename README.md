# RPA: Coleta e Consolidação de Excel (PyAutoGUI + Pandas)

Automação simples em Python que **simula um processo manual** em um portal web: abre o navegador, faz login, aplica filtros (filial + período), baixa um **relatório em Excel** e gera um **resumo consolidado por data** no terminal, salvando também um `CSV`.

> Projeto **profissional sanitizado** (antigo estágio): não contém dados reais, credenciais, nomes internos ou links privados.

---

## Por que esse projeto existe?

Na prática, eu precisava realizar tarefas repetitivas para **coletar valores/relatórios** em um portal web e depois **consolidar no Excel** para conciliação.  
Esse script foi uma forma de aprender automação (RPA) e reduzir trabalho manual.

---

## O que ele faz

1. Abre o Chrome e acessa um portal (URL configurável)
2. Realiza login (usuário/senha via variáveis de ambiente ou prompt)
3. Seleciona “filial” e intervalo de datas
4. Clica para gerar/baixar um Excel
5. Lê o arquivo baixado e gera um **resumo por data** (soma de “Quantia”)
6. Salva o resultado em `output/resumo.csv`

---

## Requisitos

- Python 3.10+ (recomendado)
- Dependências: `pyautogui`, `pandas`, `openpyxl`

Instalação:

```bash
pip install -r requirements.txt
