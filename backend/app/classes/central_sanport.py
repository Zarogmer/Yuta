from yuta_helpers import *
from .faturamento_completo import FaturamentoCompleto
from .faturamento_atipico import FaturamentoAtipico
from .faturamento_de_acordo import FaturamentoDeAcordo
from .programa_copiar_periodo import ProgramaCopiarPeriodo
from .programa_remover_periodo import ProgramaRemoverPeriodo
from .faturamento_sao_sebastiao import FaturamentoSaoSebastiao
from .gerar_relatorio import GerarRelatorio


class CentralSanport:
    def __init__(self):
        self.opcoes = [
            "FATURAMENTO",
            "FATURAMENTO SÃO SEBASTIÃO",
            "DE ACORDO",
            "FAZER PONTO",
            "DESFAZER PONTO - X",
            "RELATÓRIO - X",
            "SAIR DO PROGRAMA"
        ]

        # 🔹 instâncias (recomendo instanciar sob demanda p/ não carregar Excel antes)
        self.de_acordo = FaturamentoDeAcordo()
        self.relatorio = GerarRelatorio()

    # =========================
    # UTILITÁRIOS
    # =========================
    def limpar_tela(self):
        os.system("cls" if os.name == "nt" else "clear")

    def limpar_buffer_teclado(self):
        while msvcrt.kbhit():
            msvcrt.getch()

    def pausar_e_voltar(self, selecionado):
        print("\n🔁 Pressione ENTER para voltar ao menu...")
        while True:
            key = msvcrt.getch()
            if key in (b"\r", b"\n"):
                self.limpar_buffer_teclado()
                self.mostrar_menu(selecionado)
                return

    # =========================
    # MENU PRINCIPAL
    # =========================
    def mostrar_menu(self, selecionado):
        self.limpar_tela()

        print("╔" + "═" * 62 + "╗")
        print(f"║{' 🚢 CENTRAL DE PROCESSOS - SANPORT 🚢 '.center(60)}║")
        print("╚" + "═" * 62 + "╝\n")

        for i, opcao in enumerate(self.opcoes):
            if i == selecionado:
                print(f"          ►► {opcao} ◄◄")
            else:
                print(f"              {opcao}")

        print("\n" + "═" * 64)
        print("   ↑ ↓ = Navegar     ENTER = Selecionar")
        print("═" * 64)

    # =========================
    # SUBMENU FATURAMENTO
    # =========================
    def menu_faturamento(self):
        opcoes = [
            "Faturamento (Normal)",
            "Faturamento Atípico",
            "Voltar"
        ]
        selecionado = 0

        while True:
            self.limpar_tela()
            print("╔" + "═" * 62 + "╗")
            print(f"║{' 💰 MENU FATURAMENTO 💰 '.center(60)}║")
            print("╚" + "═" * 62 + "╝\n")

            for i, opcao in enumerate(opcoes):
                if i == selecionado:
                    print(f"          ►► {opcao} ◄◄")
                else:
                    print(f"              {opcao}")

            print("\n" + "═" * 64)
            print("   ↑ ↓ = Navegar     ENTER = Selecionar")
            print("═" * 64)

            key = msvcrt.getch()

            # setas
            if key in (b"\xe0", b"\x00"):
                key = msvcrt.getch()
                if key == b"H":
                    selecionado = max(0, selecionado - 1)
                elif key == b"P":
                    selecionado = min(len(opcoes) - 1, selecionado + 1)
                continue

            # enter
            if key in (b"\r", b"\n"):
                self.limpar_tela()

                # NORMAL
                if selecionado == 0:
                    print("╔" + "═" * 62 + "╗")
                    print("║" + " INICIANDO FATURAMENTO (NORMAL)... ".center(60) + "║")
                    print("╚" + "═" * 62 + "╝\n")

                    try:
                        FaturamentoCompleto().executar()
                    except Exception as e:
                        print(f"\n❌ ERRO NO FATURAMENTO: {e}")

                    print("\n🔁 Pressione ENTER para voltar...")
                    while msvcrt.getch() not in (b"\r", b"\n"):
                        pass

                # ATÍPICO
                elif selecionado == 1:
                    print("╔" + "═" * 62 + "╗")
                    print("║" + " INICIANDO FATURAMENTO (ATÍPICO)... ".center(60) + "║")
                    print("╚" + "═" * 62 + "╝\n")

                    try:
                        FaturamentoAtipico().executar()
                    except Exception as e:
                        print(f"\n❌ ERRO NO FATURAMENTO ATÍPICO: {e}")

                    print("\n🔁 Pressione ENTER para voltar...")
                    while msvcrt.getch() not in (b"\r", b"\n"):
                        pass

                # VOLTAR
                else:
                    return

    # =========================
    # EXECUÇÃO PRINCIPAL
    # =========================
    def rodar(self):
        selecionado = 0
        self.mostrar_menu(selecionado)

        while True:
            key = msvcrt.getch()

            # SETAS
            if key in (b"\xe0", b"\x00"):
                key = msvcrt.getch()

                if key == b"H":  # ↑
                    selecionado = max(0, selecionado - 1)
                    self.mostrar_menu(selecionado)

                elif key == b"P":  # ↓
                    selecionado = min(len(self.opcoes) - 1, selecionado + 1)
                    self.mostrar_menu(selecionado)

                continue

            # ENTER → EXECUTA A OPÇÃO
            if key in (b"\r", b"\n"):
                self.limpar_tela()

                # ----------------------------
                # FATURAMENTO (SUBMENU)
                # ----------------------------
                if selecionado == 0:
                    self.menu_faturamento()
                    self.mostrar_menu(selecionado)

                # ----------------------------
                # FATURAMENTO SÃO SEBASTIÃO
                # ----------------------------
                elif selecionado == 1:
                    print("╔" + "═" * 62 + "╗")
                    print("║" + " INICIANDO FATURAMENTO SÃO SEBASTIÃO... ".center(60) + "║")
                    print("╚" + "═" * 62 + "╝\n")

                    try:
                        programa = FaturamentoSaoSebastiao()
                        programa.executar()
                    except Exception as e:
                        print(f"\n❌ ERRO NO FATURAMENTO SSZ: {e}")

                    self.pausar_e_voltar(selecionado)

                # ----------------------------
                # DE ACORDO
                # ----------------------------
                elif selecionado == 2:
                    print("╔" + "═" * 62 + "╗")
                    print("║" + " INICIANDO DE ACORDO... ".center(60) + "║")
                    print("╚" + "═" * 62 + "╝\n")

                    try:
                        self.de_acordo.executar()
                    except Exception as e:
                        print(f"\n❌ ERRO: {e}")

                    self.pausar_e_voltar(selecionado)

                # ----------------------------
                # FAZER PONTO
                # ----------------------------
                elif selecionado == 3:
                    programa = ProgramaCopiarPeriodo(debug=True)

                    try:
                        programa.executar()
                    except Exception as e:
                        print(f"\n❌ ERRO NO FAZER PONTO: {e}")

                    self.pausar_e_voltar(selecionado)

                # ----------------------------
                # DESFAZER PONTO
                # ----------------------------
                elif selecionado == 4:
                    programa = ProgramaRemoverPeriodo(debug=True)

                    try:
                        programa.executar()
                    except Exception as e:
                        print(f"\n❌ ERRO NO DESFAZER PONTO: {e}")

                    self.pausar_e_voltar(selecionado)

                # ----------------------------
                # RELATÓRIO
                # ----------------------------
                elif selecionado == 5:
                    print("╔" + "═" * 62 + "╗")
                    print("║" + " INICIANDO RELATÓRIO... ".center(60) + "║")
                    print("╚" + "═" * 62 + "╝\n")

                    try:
                        self.relatorio.executar()
                        print("\n✅ RELATÓRIO GERADO COM SUCESSO")
                    except Exception as e:
                        print(f"\n❌ ERRO NO RELATÓRIO: {e}")

                    self.pausar_e_voltar(selecionado)

                # ----------------------------
                # SAIR
                # ----------------------------
                elif selecionado == 6:
                    self.limpar_tela()
                    print("\n👋 Saindo do programa...")
                    break
