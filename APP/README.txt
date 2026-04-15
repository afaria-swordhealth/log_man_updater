============================================
 Logistics Invoice Updater
============================================

WINDOWS
-------
1. Abre a pasta Windows
2. Duplo-clique em "install_windows.bat" (uma vez)
   - Se Python não estiver instalado, vai pedir para instalar em python.org
   - Marca "Add Python to PATH" durante a instalação
3. Duplo-clique em "Run Logistics Updater.bat" para abrir a app


MAC
---
1. Abre o Terminal (Aplicações > Utilitários > Terminal)
2. Cola este comando e pressiona Enter (só uma vez, desbloqueia os ficheiros):

   xattr -dr com.apple.quarantine ~/Desktop/APP && chmod +x ~/Desktop/APP/MAC/*.command

   (ajusta o caminho se a pasta não estiver no Desktop)

3. Duplo-clique em MAC > "install_mac.command" (uma vez)
   - Se Python não estiver instalado, abre o instalador automaticamente
   - Depois de instalar Python, corre install_mac.command outra vez
4. Duplo-clique em MAC > "Open Logistics Updater.command" para abrir a app


PROBLEMAS?
----------
Mac: corre MAC > "diagnose_mac.command" e envia o output completo.
Windows: envia screenshot do erro.
