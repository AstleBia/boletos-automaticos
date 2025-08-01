# ==============================================================================
# SCRIPT FINAL DE AUTOMAÇÃO ACADWEB - VERSÃO PROCESSADOR DE LOTES
# ==============================================================================

import time
import pyperclip
import os
from pywinauto.application import Application

# --- CONFIGURAÇÕES ---
# 1. COLOQUE AQUI TODAS AS MATRÍCULAS QUE VOCÊ QUER PROCESSAR
LISTA_MATRICULAS = [
    "123456-7",
    "987654-3",
    "112233-4",
    # Adicione quantas matrículas quiser, separadas por vírgula
]

# 2. DEFINA A PASTA ONDE OS ARQUIVOS SERÃO SALVOS
DIRETORIO_PARA_SALVAR = r"C:\boletos"

# ==============================================================================

def processar_uma_matricula(app, janela_principal, matricula):
    """
    Executa todo o fluxo de automação para uma única matrícula.
    Retorna True se bem-sucedido, False se ocorrer um erro.
    """
    try:
        print("-" * 50)
        print(f"INICIANDO PROCESSAMENTO DA MATRÍCULA: {matricula}")

        # --- AÇÃO 2: PESQUISAR A MATRÍCULA ---
        janela_principal.child_window(auto_id=2105474, control_type="Edit").set_text(matricula).type_keys("{ENTER}")
        print(f"[{matricula}] Pesquisa realizada.")

        # --- AÇÃO 3: CLICAR COM O BOTÃO DIREITO NO ALUNO ---
        resultado_aluno = janela_principal.child_window(auto_id=2951858)
        resultado_aluno.wait('visible ready', timeout=10)
        resultado_aluno.right_click_input()
        print(f"[{matricula}] Clique direito no aluno.")

        # --- AÇÃO 4: COPIAR NOME E SALVAR EM VARIÁVEL ---
        app.PopupMenu.menu_select("Copiar célula")
        time.sleep(0.5)
        valor_copiado = pyperclip.paste().strip()
        print(f"[{matricula}] Valor copiado: '{valor_copiado}'")

        # --- AÇÃO 5: NAVEGAR PARA O FINANCEIRO ---
        janela_principal.child_window(title="Financeiro").click_input()
        time.sleep(2)
        print(f"[{matricula}] Navegou para a área Financeiro.")

        # --- AÇÃO 6: SELECIONAR TODOS OS DÉBITOS ---
        painel_financeiro = janela_principal.child_window(auto_id=22808452)
        painel_financeiro.wait('visible ready', timeout=10)
        painel_financeiro.right_click_input()
        time.sleep(1)
        app.PopupMenu.child_window(title="Selecionar Todos os Débitos", auto_id=400).click_input()
        print(f"[{matricula}] Débitos selecionados.")
        time.sleep(1)
        
        # --- AÇÃO 7: IMPRIMIR BOLETOS ---
        painel_financeiro.right_click_input()
        time.sleep(1)
        app.PopupMenu.child_window(title="Imprimir Boletos Selecionados", auto_id=394).click_input()
        print(f"[{matricula}] Comando de impressão enviado.")

        # --- AÇÃO 8: EXPORTAR RELATÓRIO ---
        janela_relatorio = app.window(title="Visualização do relatório", class_name="TFRM_PreviewCrystal")
        janela_relatorio.wait('visible ready', timeout=40)
        janela_relatorio.set_focus()
        janela_relatorio.child_window(child_id=9, control_type="Button").click_input()
        print(f"[{matricula}] Botão de exportar clicado.")

        # --- AÇÃO 9: CONFIRMAR EXPORTAÇÃO ---
        janela_export_dialog = app.window(title="Export")
        janela_export_dialog.wait('visible ready', timeout=10)
        janela_export_dialog.child_window(title="OK", control_type="Button").click_input()
        print(f"[{matricula}] Confirmação de exportação OK.")

        # --- AÇÃO 10: SALVAR O ARQUIVO ---
        janela_salvar = app.window(title="Salvar como", class_name="#32770")
        janela_salvar.wait('visible ready', timeout=15)
        nome_limpo = "".join(c for c in valor_copiado if c.isalnum() or c in " ._-").rstrip()
        # Nome de arquivo mais robusto, incluindo a matrícula
        nome_arquivo = f"boleto_{matricula}_{nome_limpo}.pdf"
        caminho_completo = os.path.join(DIRETORIO_PARA_SALVAR, nome_arquivo)
        janela_salvar.child_window(title="Nome:", control_type="Edit").set_edit_text(caminho_completo)
        janela_salvar.child_window(title="Salvar", auto_id="1", control_type="Button").click_input()
        print(f"[{matricula}] Arquivo salvo em: {caminho_completo}")
        time.sleep(3)

        # --- AÇÃO 11: RESETAR A INTERFACE PARA O PRÓXIMO CICLO ---
        print(f"[{matricula}] Resetando a interface...")
        # 1. Fecha a janela de relatório
        if janela_relatorio.exists():
            janela_relatorio.close()
            print(f"[{matricula}] Janela de relatório fechada.")
        
        # 2. Clica na aba "Aluno" para voltar ao início
        aba_aluno = janela_principal.child_window(title="Aluno", child_id=1, control_type="TabItem")
        aba_aluno.click_input()
        print(f"[{matricula}] Voltou para a aba Aluno.")
        time.sleep(2) # Pausa para garantir que a aba carregou

        return True

    except Exception as e:
        print(f"!!!! ERRO AO PROCESSAR A MATRÍCULA {matricula} !!!!")
        print(f"Detalhes: {e}")
        # Tenta resetar a interface mesmo em caso de erro
        try:
            if app.window(title="Visualização do relatório").exists():
                app.window(title="Visualização do relatório").close()
            janela_principal.child_window(title="Aluno", child_id=1, control_type="TabItem").click_input()
        except Exception as reset_e:
            print(f"Falha ao tentar resetar a interface após o erro: {reset_e}")
        return False


# --- BLOCO PRINCIPAL DE EXECUÇÃO ---
if __name__ == "__main__":
    if not os.path.exists(DIRETORIO_PARA_SALVAR):
        try:
            os.makedirs(DIRETORIO_PARA_SALVAR)
            print(f"Diretório '{DIRETORIO_PARA_SALVAR}' criado com sucesso.")
        except Exception as e:
            print(f"ERRO CRÍTICO: Não foi possível criar o diretório. Verifique as permissões. Detalhes: {e}")
            exit()
            
    print("Iniciando Robô de Processamento em Lote...")
    try:
        app = Application(backend="uia").connect(path="AcadwebCursos.exe")
        janela_principal = app.window(title_re="^Acadweb Cursos.*")
        janela_principal.wait('visible ready', timeout=30)
        janela_principal.set_focus()
        print("Conexão com a aplicação estabelecida.")
    except Exception as e:
        print(f"ERRO CRÍTICO: Não foi possível conectar à aplicação. Verifique se ela está aberta. Detalhes: {e}")
        exit()

    sucessos = 0
    falhas = 0

    for matricula_atual in LISTA_MATRICULAS:
        if processar_uma_matricula(app, janela_principal, matricula_atual):
            sucessos += 1
        else:
            falhas += 1
    
    print("\n" + "="*50)
    print("PROCESSAMENTO FINALIZADO!")
    print(f"Total de matrículas processadas: {len(LISTA_MATRICULAS)}")
    print(f"  - Sucessos: {sucessos}")
    print(f"  - Falhas: {falhas}")
    print("="*50)
