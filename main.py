import time
import pyperclip
import os
from pywinauto.application import Application


codigo_matricula_a_pesquisar = "123456-7"  # Matrícula para teste
diretorio_para_salvar = r"C:\boletos"

if not os.path.exists(diretorio_para_salvar):
    try:
        os.makedirs(diretorio_para_salvar)
        print(f"Diretório '{diretorio_para_salvar}' criado com sucesso.")
    except Exception as e:
        print(
            f"ERRO CRÍTICO: Não foi possível criar o diretório '{diretorio_para_salvar}'. Verifique as permissões. Detalhes: {e}")
        exit()

print("Conectando à aplicação 'AcadwebCursos.exe' que já está aberta...")
try:
    app = Application(backend="uia").connect(path="AcadwebCursos.exe")
    janela_principal = app.window(title_re="^Acadweb Cursos.*")
    janela_principal.wait('visible ready', timeout=30)
    janela_principal.set_focus()
    print("Conexão e janela principal OK.")
except Exception as e:
    print(f"ERRO CRÍTICO: Não foi possível conectar à aplicação. Verifique se ela está aberta. Detalhes: {e}")
    exit()

print(f"Pesquisando a matrícula: {codigo_matricula_a_pesquisar}")
try:
    barra_pesquisa = janela_principal.child_window(auto_id=2105474, control_type="Edit")
    barra_pesquisa.set_edit_text(codigo_matricula_a_pesquisar).type_keys("{ENTER}")
    print("Pesquisa realizada.")
except Exception as e:
    print(f"ERRO na Ação 2 (Pesquisar Matrícula). Detalhes: {e}")
    exit()

print("Clicando com o botão direito no resultado do aluno...")
try:
    resultado_aluno = janela_principal.child_window(auto_id=2951858)
    resultado_aluno.wait('visible ready', timeout=10)
    resultado_aluno.right_click_input()
    print("Clique direito executado.")
except Exception as e:
    print(f"ERRO na Ação 3 (Clique Direito). Detalhes: {e}")
    exit()

print("Copiando o nome da célula...")
try:
    app.PopupMenu.menu_select("Copiar célula")
    time.sleep(0.5)
    valor_copiado = pyperclip.paste().strip()
    print(f"--- VALOR COPIADO: '{valor_copiado}' ---")
except Exception as e:
    print(f"ERRO na Ação 4 (Copiar Célula). Detalhes: {e}")
    exit()

print("Navegando para a área 'Financeiro'...")
try:
    janela_principal.child_window(title="Financeiro").click_input()
    time.sleep(2)
    print("Área 'Financeiro' selecionada.")
except Exception as e:
    print(f"ERRO na Ação 5 (Navegar Financeiro). Detalhes: {e}")
    exit()

print("Selecionando todos os débitos...")
try:
    painel_financeiro = janela_principal.child_window(auto_id=22808452)
    painel_financeiro.wait('visible ready', timeout=10)
    painel_financeiro.right_click_input()
    time.sleep(1)
    app.PopupMenu.child_window(title="Selecionar Todos os Débitos", auto_id=400).click_input()
    print("Débitos selecionados.")
    time.sleep(1)
except Exception as e:
    print(f"ERRO na Ação 6 (Selecionar Débitos). Detalhes: {e}")
    exit()

print("Enviando comando para imprimir boletos...")
try:
    janela_principal.child_window(auto_id=22808452).right_click_input()
    time.sleep(1)
    app.PopupMenu.child_window(title="Imprimir Boletos Selecionados", auto_id=394).click_input()
    print("Comando de impressão enviado.")
except Exception as e:
    print(f"ERRO na Ação 7 (Imprimir Boletos). Detalhes: {e}")
    exit()


print("Aguardando janela de relatório e clicando em exportar...")
try:
    janela_relatorio = app.window(title="Visualização do relatório", class_name="TFRM_PreviewCrystal")
    janela_relatorio.wait('visible ready', timeout=40)
    janela_relatorio.set_focus()
    janela_relatorio.child_window(child_id=9, control_type="Button").click_input()
    print("Botão de exportar clicado.")
except Exception as e:
    print(f"ERRO na Ação 8 (Exportar Relatório). Detalhes: {e}")
    exit()


print("Confirmando exportação...")
try:
    janela_export_dialog = app.window(title="Export")
    janela_export_dialog.wait('visible ready', timeout=10)
    janela_export_dialog.child_window(title="OK", control_type="Button").click_input()
    print("Confirmação OK.")
except Exception as e:
    print(f"ERRO na Ação 9 (Confirmar Export). Detalhes: {e}")
    exit()


print("Salvando o arquivo final...")
try:
    # Usa os valores padrão para a janela "Salvar como". Se falhar aqui, verifique estes valores com o Inspect.exe
    janela_salvar = app.window(title="Salvar como", class_name="#32770")
    janela_salvar.wait('visible ready', timeout=15)
    print("Janela 'Salvar como' encontrada.")

    # Cria um nome de arquivo limpo a partir do valor copiado (ex: "boleto_Fulano_de_Tal.pdf")
    nome_limpo = "".join(c for c in valor_copiado if c.isalnum() or c in " ._-").rstrip()
    nome_arquivo = f"boleto_{nome_limpo}.pdf"

    # Junta o diretório e o nome do arquivo para formar o caminho completo
    caminho_completo = os.path.join(diretorio_para_salvar, nome_arquivo)
    print(f"Salvando arquivo em: {caminho_completo}")

    # Insere o caminho completo no campo de nome de arquivo
    # Usando valores padrão para o campo "Nome:". Se falhar, verifique com Inspect.exe
    campo_nome_arquivo = janela_salvar.child_window(title="Nome:", control_type="Edit")
    campo_nome_arquivo.set_edit_text(caminho_completo)

    # Clica no botão "Salvar" usando o AutomationId que você forneceu
    botao_salvar = janela_salvar.child_window(title="Salvar", auto_id="1", control_type="Button")
    botao_salvar.click_input()
    print("Botão 'Salvar' clicado.")

    # Pausa final para garantir que o arquivo foi salvo antes de terminar
    time.sleep(3)

except Exception as e:
    print(f"ERRO na Ação 10 (Salvar Arquivo).")
    print("Verifique os identificadores da janela 'Salvar como', do campo 'Nome:' e do botão 'Salvar'.")
    print(f"Detalhes do erro: {e}")
    exit()

print("\n==============================================")
print("ROBÔ FINALIZADO COM SUCESSO!")
print(f"O arquivo deve ter sido salvo em '{diretorio_para_salvar}'.")
print("==============================================")
