============================================
 Logistics Management Updater  V0.1
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
IMPORTANTE: A pasta Log_Man_Updater tem de estar no Desktop antes de começar.

1. Abre o Terminal (Aplicações > Utilitários > Terminal)
2. Desbloqueia os ficheiros (só uma vez):

   Clica 3x na linha abaixo para selecionar → Cmd+C → Terminal → Cmd+V → Enter

   xattr -dr com.apple.quarantine ~/Desktop/Log_Man_Updater && chmod +x ~/Desktop/Log_Man_Updater/MAC/*.command

3. Duplo-clique em MAC > "install_mac.command" (uma vez)
   - Se Python não estiver instalado, abre o instalador automaticamente
   - Depois de instalar Python, corre install_mac.command outra vez
4. Duplo-clique em MAC > "Open Logistics Updater.command" para abrir a app


PROBLEMAS?
----------
Mac: corre MAC > "diagnose_mac.command" e envia o output completo.
Windows: envia screenshot do erro.
