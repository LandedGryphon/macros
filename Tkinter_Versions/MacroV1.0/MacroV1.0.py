"""
Script de Automação de Mouse (Macro)
Desenvolvido com tkinter para GUI e pynput para controle do mouse
Autor: Senior Python Developer
Data: 2026-01-05

Funcionalidades:
- Interface gráfica para configuração de macros
- Captura de coordenadas da tela
- Controle de cliques do mouse (esquerdo/direito)
- Hotkeys para controle de execução (F1, F2, F3)
- Execução em thread separada para não travar a GUI
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
try:
    from pynput.mouse import Controller, Button, Listener as MouseListener
    from pynput.keyboard import Listener, Key
except ImportError:
    print("Erro: pynput não está instalado. Execute: pip install pynput")
    exit(1)


class MacroAutomation:
    """
    Classe principal para gerenciar automação de mouse com GUI.
    Responsável pela interface, captura de coordenadas e execução do macro.
    """
    
    def __init__(self, root):
        """
        Inicializa a aplicação de automação de macro.
        
        Args:
            root (tk.Tk): Janela raiz do tkinter
        """
        self.root = root
        self.root.title("Macro Automation - F1 (Start) F2 (Pause) F3 (Exit)")
        self.root.geometry("600x550")
        self.root.resizable(True, True)  # Agora é redimensionável
        self.root.minsize(400, 350)  # Tamanho mínimo
        
        # Variáveis de controle
        self.mouse_controller = Controller()  # Controlador do mouse
        self.is_running = False  # Status da execução do macro
        self.is_paused = False   # Status de pausa
        self.capture_mode = False  # Modo de captura de coordenadas
        
        # Coordenadas salvas
        self.saved_x = None
        self.saved_y = None
        
        # Tipo de botão selecionado
        self.button_type = tk.StringVar(value="esquerdo")
        
        # Thread para execução do macro
        self.macro_thread = None
        
        # Listeners
        self.listener = None
        self.mouse_listener = None
        
        # Keybinds configuráveis
        self.key_start = Key.f1
        self.key_pause = Key.f2
        self.key_exit = Key.f3
        
        # Inicializar listeners de forma segura
        self._initialize_listeners()
        
        # Construir interface gráfica
        self._build_gui()
        
        # Gerenciar fechamento da janela
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _initialize_listeners(self):
        """Inicializa os listeners de teclado e mouse de forma segura."""
        try:
            self.mouse_listener = MouseListener(on_click=self._on_mouse_click)
            self.mouse_listener.start()
            
            self.listener = Listener(on_press=self._on_key_press)
            self.listener.start()
        except Exception as e:
            print(f"Erro ao inicializar listeners: {str(e)}")
    
    def _build_gui(self):
        """Constrói a interface gráfica com todos os widgets."""
        
        # Criar menu bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # Menu Arquivo
        arquivo_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Arquivo", menu=arquivo_menu)
        arquivo_menu.add_command(label="Sair", command=self._on_closing)
        
        # Menu Configurações
        config_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Configurações", menu=config_menu)
        config_menu.add_command(label="Rebindar Teclas...", command=self._open_keybind_dialog)
        config_menu.add_separator()
        config_menu.add_command(label="Restaurar Padrão", command=self._reset_keybinds)
        
        # Menu Ajuda
        ajuda_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ajuda", menu=ajuda_menu)
        ajuda_menu.add_command(label="Sobre", command=self._show_about)
        
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ===== SEÇÃO: Seleção de Botão =====
        button_frame = ttk.LabelFrame(main_frame, text="Seleção de Botão", padding="10")
        button_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Radiobutton(
            button_frame,
            text="Botão Esquerdo",
            variable=self.button_type,
            value="esquerdo"
        ).pack(anchor=tk.W)
        
        ttk.Radiobutton(
            button_frame,
            text="Botão Direito",
            variable=self.button_type,
            value="direito"
        ).pack(anchor=tk.W)
        
        # ===== SEÇÃO: Captura de Coordenadas =====
        coord_frame = ttk.LabelFrame(main_frame, text="Definição de Local", padding="10")
        coord_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Label para exibir coordenadas capturadas
        self.coord_label = ttk.Label(
            coord_frame,
            text="Coordenadas: Não capturadas",
            font=("Arial", 10),
            foreground="gray"
        )
        self.coord_label.pack(pady=(0, 10))
        
        # Botão para capturar coordenadas
        self.capture_button = ttk.Button(
            coord_frame,
            text="Capturar Coordenada",
            command=self._start_capture_mode
        )
        self.capture_button.pack(fill=tk.X)
        
        # ===== SEÇÃO: Instruções de Hotkeys =====
        hotkey_frame = ttk.LabelFrame(main_frame, text="Hotkeys de Controle", padding="10")
        hotkey_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        hotkey_info = (
            f"{self._key_name(self.key_start)} - Iniciar (Start)\n"
            f"{self._key_name(self.key_pause)} - Pausar/Parar (Pause/Stop)\n"
            f"{self._key_name(self.key_exit)} - Sair (Exit)\n\n"
            "Dica: Configure a coordenada antes de iniciar!\n"
            "A janela pode ser minimizada durante a execução.\n"
            "Acesse Configurações > Rebindar Teclas para mudar os hotkeys."
        )
        
        self.hotkey_label = ttk.Label(
            hotkey_frame,
            text=hotkey_info,
            font=("Arial", 9),
            justify=tk.LEFT
        )
        self.hotkey_label.pack(anchor=tk.W)
        
        # ===== SEÇÃO: Status =====
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding="10")
        status_frame.pack(fill=tk.X)
        
        self.status_label = ttk.Label(
            status_frame,
            text="Status: Pronto",
            font=("Arial", 9),
            foreground="green"
        )
        self.status_label.pack(anchor=tk.W)
    
    def _start_capture_mode(self):
        """
        Ativa o modo de captura de coordenadas.
        O usuário pode clicar em qualquer ponto da tela para capturar as coordenadas.
        """
        # Mostrar instrução ANTES de ativar o capture mode
        messagebox.showinfo(
            "Captura de Coordenadas",
            "Clique em qualquer ponto da tela para capturar a coordenada.\n"
            "O capture mode será desativado automaticamente após o clique."
        )
        
        # DEPOIS que o usuário clica OK, ativar o capture mode
        self.capture_mode = True
        self.capture_button.config(state=tk.DISABLED, text="Aguardando clique...")
        self.status_label.config(text="Status: Aguardando clique na tela...", foreground="orange")
        self.root.update()
    
    def _on_key_press(self, key):
        """
        Manipulador de eventos de teclado para hotkeys.
        
        Args:
            key: Tecla pressionada
        """
        try:
            # F1 - Iniciar (ou tecla configurada)
            if key == self.key_start:
                if self.saved_x is None or self.saved_y is None:
                    messagebox.showwarning(
                        "Aviso",
                        "Por favor, capture uma coordenada primeiro!"
                    )
                    return
                
                self.is_running = True
                self.is_paused = False
                self._update_status("Executando...", "green")
                
                # Iniciar thread para execução
                self.macro_thread = threading.Thread(target=self._execute_macro, daemon=True)
                self.macro_thread.start()
            
            # F2 - Pausar/Parar (ou tecla configurada)
            elif key == self.key_pause:
                self.is_running = False
                self.is_paused = True
                self._release_mouse()
                self._update_status("Pausado", "orange")
            
            # F3 - Sair (ou tecla configurada)
            elif key == self.key_exit:
                self._on_closing()
        
        except AttributeError:
            # Teclas comuns que não são atributos de Key
            pass
    
    def _execute_macro(self):
        """
        Executa a automação do macro em thread separada.
        Move o mouse para a coordenada e mantém o botão pressionado.
        """
        try:
            # Determinar qual botão será usado
            button = Button.left if self.button_type.get() == "esquerdo" else Button.right
            
            # Mover mouse para coordenada salva
            self.mouse_controller.position = (self.saved_x, self.saved_y)
            time.sleep(0.1)  # Pequeno delay para garantir que o mouse chegou
            
            # Pressionar e manter o botão
            self.mouse_controller.press(button)
            
            # Manter o botão pressionado enquanto is_running for True
            while self.is_running:
                time.sleep(0.05)  # Check a cada 50ms
            
            # Soltar o botão quando parar
            self.mouse_controller.release(button)
            self._update_status("Parado", "red")
        
        except Exception as e:
            self._update_status(f"Erro: {str(e)}", "red")
            print(f"Erro durante execução: {str(e)}")
    
    def _release_mouse(self):
        """Solta qualquer botão do mouse que esteja pressionado."""
        try:
            # Tentar soltar ambos os botões para garantir
            self.mouse_controller.release(Button.left)
        except:
            pass
        
        try:
            self.mouse_controller.release(Button.right)
        except:
            pass
    
    def _update_status(self, message, color):
        """
        Atualiza o label de status na GUI.
        
        Args:
            message (str): Mensagem de status
            color (str): Cor do texto
        """
        try:
            self.status_label.config(text=f"Status: {message}", foreground=color)
            self.root.update()
        except:
            pass  # Evitar erros se a janela foi fechada
    
    def _on_mouse_click(self, x, y, button, pressed):
        """
        Manipulador de eventos de mouse para captura de coordenadas.
        
        Args:
            x (int): Coordenada X do clique
            y (int): Coordenada Y do clique
            button: Botão do mouse
            pressed (bool): True se pressionado, False se solto
        """
        # Apenas capturar na pressão do botão
        if self.capture_mode and pressed:
            self.saved_x = x
            self.saved_y = y
            self.capture_mode = False
            
            # Atualizar GUI
            self.coord_label.config(
                text=f"Coordenadas: X={self.saved_x}, Y={self.saved_y}",
                foreground="green"
            )
            self.capture_button.config(state=tk.NORMAL, text="Capturar Coordenada")
            self._update_status("Coordenada capturada!", "green")
    
    def _key_name(self, key):
        """Retorna o nome legível da tecla."""
        try:
            return key.name.upper()
        except:
            return str(key).upper()
    
    def _open_keybind_dialog(self):
        """Abre um diálogo para rebindar as teclas."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Rebindar Teclas")
        dialog.geometry("400x300")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Frame principal
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Pressione a tecla desejada para cada ação:", font=("Arial", 10)).pack(pady=(0, 20))
        
        # Tecla de Start
        start_frame = ttk.Frame(main_frame)
        start_frame.pack(fill=tk.X, pady=10)
        ttk.Label(start_frame, text="Iniciar (Start):", width=20).pack(side=tk.LEFT)
        start_button = ttk.Button(start_frame, text=self._key_name(self.key_start), width=15)
        start_button.pack(side=tk.LEFT, padx=10)
        start_button.config(command=lambda: self._capture_key("start", start_button))
        
        # Tecla de Pause
        pause_frame = ttk.Frame(main_frame)
        pause_frame.pack(fill=tk.X, pady=10)
        ttk.Label(pause_frame, text="Pausar/Parar (Pause):", width=20).pack(side=tk.LEFT)
        pause_button = ttk.Button(pause_frame, text=self._key_name(self.key_pause), width=15)
        pause_button.pack(side=tk.LEFT, padx=10)
        pause_button.config(command=lambda: self._capture_key("pause", pause_button))
        
        # Tecla de Exit
        exit_frame = ttk.Frame(main_frame)
        exit_frame.pack(fill=tk.X, pady=10)
        ttk.Label(exit_frame, text="Sair (Exit):", width=20).pack(side=tk.LEFT)
        exit_button = ttk.Button(exit_frame, text=self._key_name(self.key_exit), width=15)
        exit_button.pack(side=tk.LEFT, padx=10)
        exit_button.config(command=lambda: self._capture_key("exit", exit_button))
        
        # Botões
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        ttk.Button(button_frame, text="OK", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Restaurar Padrão", command=self._reset_keybinds).pack(side=tk.LEFT, padx=5)
    
    def _capture_key(self, action, button):
        """Captura uma tecla pressionada em thread separada."""
        button.config(text="Aguardando tecla...", state=tk.DISABLED)
        self.root.update()
        
        def capture_in_thread():
            """Executa a captura em thread separada para não travar a GUI."""
            def on_press(key):
                try:
                    if action == "start":
                        self.key_start = key
                    elif action == "pause":
                        self.key_pause = key
                    elif action == "exit":
                        self.key_exit = key
                    
                    # Atualizar GUI de forma segura
                    self.root.after(0, lambda: button.config(text=self._key_name(key), state=tk.NORMAL))
                    self.root.after(0, self._update_hotkey_display)
                    return False  # Parar o listener
                except:
                    return False
            
            try:
                with Listener(on_press=on_press) as listener:
                    listener.join(timeout=5)  # Esperar até 5 segundos
            except:
                pass
            
            # Restaurar botão se nenhuma tecla foi pressionada
            self.root.after(0, lambda: button.config(state=tk.NORMAL))
        
        # Executar captura em thread separada
        capture_thread = threading.Thread(target=capture_in_thread, daemon=True)
        capture_thread.start()
    
    def _reset_keybinds(self):
        """Restaura as teclas padrão."""
        self.key_start = Key.f1
        self.key_pause = Key.f2
        self.key_exit = Key.f3
        self._update_hotkey_display()
        messagebox.showinfo("Sucesso", "Teclas restauradas ao padrão!")
    
    def _update_hotkey_display(self):
        """Atualiza a exibição dos hotkeys na GUI."""
        try:
            hotkey_info = (
                f"{self._key_name(self.key_start)} - Iniciar (Start)\n"
                f"{self._key_name(self.key_pause)} - Pausar/Parar (Pause/Stop)\n"
                f"{self._key_name(self.key_exit)} - Sair (Exit)\n\n"
                "Dica: Configure a coordenada antes de iniciar!\n"
                "A janela pode ser minimizada durante a execução.\n"
                "Acesse Configurações > Rebindar Teclas para mudar os hotkeys."
            )
            self.hotkey_label.config(text=hotkey_info)
            self.root.update()
        except:
            pass
    
    def _show_about(self):
        """Mostra a janela de sobre."""
        about_text = """Macro Automation v1.0

Automação profissional de mouse com GUI

Desenvolvido em Python com:
- tkinter (Interface Gráfica)
- pynput (Controle do Mouse)

Autor: Senior Python Developer
Data: 2026-01-05

Features:
✓ Interface redimensionável
✓ Captura de coordenadas
✓ Hotkeys configuráveis
✓ Execução em thread separada
✓ Controle robusto do mouse

Hotkeys Padrão:
F1 - Iniciar
F2 - Pausar
F3 - Sair
"""
        messagebox.showinfo("Sobre", about_text)
    
    def _on_closing(self):
        """Encerra a aplicação de forma segura, liberando recursos."""
        # Parar execução
        self.is_running = False
        
        # Soltar mouse se estiver pressionado
        self._release_mouse()
        
        # Parar listeners
        try:
            if self.listener:
                self.listener.stop()
        except:
            pass
        
        try:
            if self.mouse_listener:
                self.mouse_listener.stop()
        except:
            pass
        
        # Fechar janela
        self.root.destroy()


def main():
    """Função principal para inicializar a aplicação."""
    try:
        # Criar janela raiz
        root = tk.Tk()
        
        # Instanciar aplicação
        app = MacroAutomation(root)
        
        # Iniciar loop principal da GUI
        root.mainloop()
    
    except Exception as e:
        print(f"Erro ao inicializar aplicação: {str(e)}")
        exit(1)


if __name__ == "__main__":
    main()
