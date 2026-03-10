from backend.app.modules.gerador_nf import *


if __name__ == "__main__":
    import sys
    from PyQt6.QtWidgets import QApplication

    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    w = GeradorNFSe()
    w.show()
    sys.exit(app.exec())
