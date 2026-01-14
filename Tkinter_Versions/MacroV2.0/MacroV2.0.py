"""
Script de Automação de Mouse e Teclado (Macro) - V2.0
Desenvolvido com tkinter para GUI e pynput para controle
Autor: Senior Python Developer
Data: 2026-01-06

Funcionalidades V2.0:
- Interface gráfica moderna com tema claro/escuro
- Captura de coordenadas da tela
- Controle de cliques do mouse (esquerdo/direito) com press and hold
- Controle de teclado (press and hold)
- Hotkeys para controle de execução (F1, F2, F3)
- Configuração de delay entre cliques em milissegundos
- Execução em thread separada para não travar a GUI
- Salvamento automático de configurações
- Sistema de temas (Dark/Light)
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import threading
import time
import json
import os
import sys
from pathlib import Path
try:
    from pynput.mouse import Controller as MouseController, Button, Listener as MouseListener
    from pynput.keyboard import Controller as KeyboardController, Listener, Key
except ImportError:
    print("Erro: pynput não está instalado. Execute: pip install pynput")
    exit(1)

# ===== VARIÁVEL GLOBAL CRÍTICA =====
# Capturar o diretório correto ANTES de qualquer mudança
# Isso é essencial para PyInstaller --onefile
if getattr(sys, 'frozen', False):
    # Rodando como .exe - usar o módulo inspect para pegar o local real
    import inspect
    _BASE_DIR = Path(inspect.getfile(lambda: None)).resolve().parent
    if "temp" in str(_BASE_DIR).lower():
        # Se ainda está em temp, usar o home do usuário
        _BASE_DIR = Path.home()
else:
    # Script Python normal
    _BASE_DIR = Path(__file__).resolve().parent

print(f"=== Iniciando em {_BASE_DIR} ===")


class ConfigManager:
    """Gerencia salvamento e carregamento de configurações."""
    
    def __init__(self, config_file="macro_config.json"):
        """Inicializa o gerenciador de configurações."""
        config_path = Path(config_file)
        if not config_path.is_absolute():
            # Determinar diretório base - importante para .exe compilados
            base_dir = None
            
            # Se está rodando como .exe congelado (PyInstaller)
            if getattr(sys, 'frozen', False):
                # Em PyInstaller --onefile, sempre usar o diretório de trabalho atual
                # pois quando o .exe é executado, o cwd é onde ele está localizado
                base_dir = Path.cwd()
                print(f"DEBUG: .exe congelado, usando cwd: {base_dir}")
            else:
                # Rodando como script Python normal
                base_dir = Path(__file__).resolve().parent
                print(f"DEBUG: Script Python normal - __file__={__file__}")
            
            # Tentar salvar no diretório base
            test_file = base_dir / ".write_test"
            try:
                test_file.write_text("test")
                test_file.unlink()
                config_path = base_dir / config_file
                print(f"DEBUG: ✓ Conseguiu escrever em {base_dir}")
            except (PermissionError, OSError) as e:
                # Fallback para diretório de documentos do usuário
                print(f"DEBUG: ✗ Não conseguiu escrever em {base_dir}: {e}")
                docs_dir = Path.home() / "Documents"
                if not docs_dir.exists():
                    docs_dir = Path.home()
                config_path = docs_dir / config_file
                print(f"DEBUG: Usando fallback: {config_path}")
                
        self.config_file = config_path
        self.config_file.parent.mkdir(parents=True, exist_ok=True)
        print(f"Arquivo de config: {self.config_file}")
        self.default_config = {
            "theme": "dark",
            "button_type": "esquerdo",
            "saved_x": None,
            "saved_y": None,
            "key_start": "f1",
            "key_pause": "f2",
            "key_exit": "f3",
            "click_delay_ms": 100,
            "action_type": "click",  # 'click' ou 'hold'
            "hold_duration_ms": 500,
            "custom_key_name": "Nenhuma"
        }
        self.config = self.load_config()
    
    def load_config(self):
        """Carrega configurações do arquivo JSON."""
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # Mesclar com padrão para garantir que todas as chaves existem
                    return {**self.default_config, **config}
            except Exception as e:
                print(f"Erro ao carregar configurações: {e}")
                return self.default_config.copy()
        return self.default_config.copy()
    
    def save_config(self):
        """Salva configurações no arquivo JSON."""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Erro ao salvar configurações: {e}")
            return False
    
    def get(self, key, default=None):
        """Obtém valor da configuração."""
        return self.config.get(key, default)
    
    def set(self, key, value):
        """Define valor de configuração e salva."""
        self.config[key] = value
        self.save_config()


class ThemeManager:
    """Gerencia temas da aplicação."""
    
    THEMES = {
        "dark": {
            "bg": "#1e1e1e",
            "fg": "#ffffff",
            "frame_bg": "#2d2d2d",
            "button_bg": "#404040",
            "button_fg": "#ffffff",
            "accent": "#0078d4",
            "success": "#4ec9b0",
            "warning": "#dcdcaa",
            "error": "#f48771"
        },
        "light": {
            "bg": "#ffffff",
            "fg": "#000000",
            "frame_bg": "#f0f0f0",
            "button_bg": "#e0e0e0",
            "button_fg": "#000000",
            "accent": "#0078d4",
            "success": "#008000",
            "warning": "#ff8c00",
            "error": "#ff0000"
        }
    }
    
    @classmethod
    def get_theme(cls, theme_name):
        """Retorna dicionário de cores do tema."""
        return cls.THEMES.get(theme_name, cls.THEMES["dark"])
    
    @classmethod
    def configure_style(cls, style, theme_name):
        """Configura estilo do ttk para o tema."""
        theme = cls.get_theme(theme_name)
        
        style.theme_use('clam')
        
        # Configurar cores do tema
        style.configure("TFrame", background=theme["frame_bg"], foreground=theme["fg"])
        style.configure("TLabel", background=theme["frame_bg"], foreground=theme["fg"])
        style.configure("TButton", background=theme["button_bg"], foreground=theme["button_fg"])
        style.configure("TLabelFrame", background=theme["frame_bg"], foreground=theme["fg"])
        style.configure("TLabelFrame.Label", background=theme["frame_bg"], foreground=theme["fg"])
        style.configure("TRadiobutton", background=theme["frame_bg"], foreground=theme["fg"])
        style.configure("TCheckbutton", background=theme["frame_bg"], foreground=theme["fg"])
        
        # Configurar scrollbar
        style.configure("Vertical.TScrollbar", background=theme["button_bg"])


class MacroAutomation:
    """
    Classe principal para gerenciar automação de mouse e teclado com GUI.
    """
    
    def __init__(self, root):
        """Inicializa a aplicação de automação de macro V2.0."""
        self.root = root
        self.root.title("Macro Automation V2.0 - F1 (Start) F2 (Pause) F3 (Exit)")
        self.root.geometry("700x700")
        self.root.resizable(True, True)
        self.root.minsize(500, 600)
        
        # Gerenciador de configurações
        # Tentar salvar na mesma pasta do script/exe
        script_dir = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(script_dir, "macro_config.json")
        
        # Se não conseguir escrever na pasta do script, usar a pasta atual
        try:
            os.makedirs(script_dir, exist_ok=True)
            # Testar se consegue escrever
            test_file = os.path.join(script_dir, ".write_test")
            with open(test_file, "w") as f:
                f.write("test")
            os.remove(test_file)
        except:
            # Usar pasta atual se não conseguir escrever na pasta do script
            config_path = os.path.join(os.getcwd(), "macro_config.json")
        
        self.config_mgr = ConfigManager(config_path)
        
        log_msg = f"Arquivo de config: {self.config_mgr.config_file}\n"
        print(log_msg, end="")
        try:
            with open(os.path.join(os.getcwd(), "macro_debug.log"), "a") as f:
                f.write(log_msg)
        except:
            pass
        
        # Tema atual
        self.current_theme = self.config_mgr.get("theme", "dark")
        self.theme = ThemeManager.get_theme(self.current_theme)
        
        # Configurar estilo
        self.style = ttk.Style()
        ThemeManager.configure_style(self.style, self.current_theme)
        
        # Aplicar cores de fundo
        self.root.configure(bg=self.theme["bg"])
        
        # Variáveis de controle
        self.mouse_controller = MouseController()
        self.keyboard_controller = KeyboardController()
        self.is_running = False
        self.is_paused = False
        self.capture_mode = False
        
        # Coordenadas salvas
        self.saved_x = self.config_mgr.get("saved_x")
        self.saved_y = self.config_mgr.get("saved_y")
        
        # Tipo de botão selecionado
        self.button_type = tk.StringVar(value=self.config_mgr.get("button_type", "esquerdo"))
        
        # Tecla customizada (para suportar qualquer tecla)
        self.custom_key = self.config_mgr.get("custom_key", None)
        self.custom_key_name = self.config_mgr.get("custom_key_name", "Nenhuma")
        
        # Tipo de ação (clique ou pressão prolongada)
        self.action_type = tk.StringVar(value=self.config_mgr.get("action_type", "click"))
        
        # Delay entre cliques (em milissegundos)
        self.click_delay_ms = tk.IntVar(value=self.config_mgr.get("click_delay_ms", 100))
        
        # Duração de pressão prolongada (em milissegundos)
        self.hold_duration_ms = tk.IntVar(value=self.config_mgr.get("hold_duration_ms", 500))
        
        # Thread para execução do macro
        self.macro_thread = None
        
        # Listeners
        self.listener = None
        self.mouse_listener = None
        
        # Keybinds
        key_start_str = self.config_mgr.get("key_start", "f1")
        key_pause_str = self.config_mgr.get("key_pause", "f2")
        key_exit_str = self.config_mgr.get("key_exit", "f3")
        log_msg = f"Carregando hotkeys: START={key_start_str}, PAUSE={key_pause_str}, EXIT={key_exit_str}\n"
        print(log_msg, end="")
        try:
            with open(os.path.join(os.getcwd(), "macro_debug.log"), "a") as f:
                f.write(log_msg)
        except:
            pass
        
        self.key_start = self._string_to_key(key_start_str)
        self.key_pause = self._string_to_key(key_pause_str)
        self.key_exit = self._string_to_key(key_exit_str)
        
        log_msg = f"Hotkeys convertidos: START={self.key_start}, PAUSE={self.key_pause}, EXIT={self.key_exit}\n"
        print(log_msg, end="")
        try:
            with open(os.path.join(os.getcwd(), "macro_debug.log"), "a") as f:
                f.write(log_msg)
        except:
            pass
        
        # Inicializar listeners
        self._initialize_listeners()
        
        # Construir interface gráfica
        self._build_gui()
        
        # Gerenciar fechamento da janela
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _string_to_key(self, key_str):
        """Converte string para objeto Key do pynput."""
        try:
            key_str_clean = key_str.lower().strip()
            # Tentar com underscore
            try:
                result = getattr(Key, key_str_clean)
                print(f"_string_to_key: {key_str} -> Key.{key_str_clean} (sucesso sem underscore)")
                return result
            except AttributeError as e1:
                # Tentar com underscore adicionado (para caracteres simples como 'a' -> '_a')
                try:
                    result = getattr(Key, f"_{key_str_clean}")
                    print(f"_string_to_key: {key_str} -> Key._{key_str_clean} (sucesso com underscore)")
                    return result
                except AttributeError:
                    # Se é um caractere simples, retornar como está
                    if len(key_str_clean) == 1:
                        print(f"_string_to_key: {key_str} -> KeyCode simples (caractere único)")
                        # Retornar um objeto que pode ser usado com keyboard_controller
                        # pynput suporta strings simples diretamente
                        return key_str_clean
                    raise
        except Exception as e:
            print(f"_string_to_key erro: {key_str} -> retornando Key.f1 (erro: {e})")
            return Key.f1  # Padrão
    
    def _key_to_string(self, key):
        """Converte Key para string."""
        try:
            key_name = key.name.lower()
            if key_name.startswith("_"):
                key_name = key_name[1:]
            print(f"_key_to_string: {key} -> name={key.name} -> final={key_name}")
            return key_name
        except Exception as e:
            print(f"Erro em _key_to_string: {e}, retornando f1")
            return "f1"
    
    def _normalize_key(self, key):
        """Normaliza uma tecla para comparação."""
        try:
            # Se é uma string, retornar como está
            if isinstance(key, str):
                return key.lower()
            # Se é um Key enum, retornar o nome
            try:
                return key.name.lower().lstrip("_")
            except:
                return str(key).lower()
        except:
            return str(key).lower()
    
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
        """Constrói a interface gráfica completa."""
        
        # Criar menu bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # Menu Arquivo
        arquivo_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Arquivo", menu=arquivo_menu)
        arquivo_menu.add_command(label="Salvar Config Manualmente", command=self._save_config_manually)
        arquivo_menu.add_command(label="Criar Config Padrão", command=self._create_default_config)
        arquivo_menu.add_separator()
        arquivo_menu.add_command(label="Sair", command=self._on_closing)
        
        # Menu Configurações
        config_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Configurações", menu=config_menu)
        config_menu.add_command(label="Rebindar Teclas...", command=self._open_keybind_dialog)
        config_menu.add_command(label="Alterar Tema...", command=self._open_theme_dialog)
        config_menu.add_separator()
        config_menu.add_command(label="Restaurar Padrão", command=self._reset_all)
        
        # Menu Ajuda
        ajuda_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ajuda", menu=ajuda_menu)
        ajuda_menu.add_command(label="Sobre", command=self._show_about)
        
        # Frame principal com scrolling
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ===== SEÇÃO: Tipo de Ação =====
        action_frame = ttk.LabelFrame(main_frame, text="Tipo de Ação", padding="10")
        action_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Radiobutton(
            action_frame,
            text="Clique Único",
            variable=self.action_type,
            value="click",
            command=self._on_action_change
        ).pack(anchor=tk.W)
        
        ttk.Radiobutton(
            action_frame,
            text="Pressionar e Manter (Hold)",
            variable=self.action_type,
            value="hold",
            command=self._on_action_change
        ).pack(anchor=tk.W)
        
        # ===== SEÇÃO: Seleção de Botão =====
        button_frame = ttk.LabelFrame(main_frame, text="Seleção de Botão/Tecla", padding="10")
        button_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Radiobutton(
            button_frame,
            text="Botão Esquerdo do Mouse",
            variable=self.button_type,
            value="esquerdo"
        ).pack(anchor=tk.W)
        
        ttk.Radiobutton(
            button_frame,
            text="Botão Direito do Mouse",
            variable=self.button_type,
            value="direito"
        ).pack(anchor=tk.W)
        
        ttk.Radiobutton(
            button_frame,
            text="Qualquer Tecla (Customizada)",
            variable=self.button_type,
            value="custom"
        ).pack(anchor=tk.W)
        
        # Label para exibir tecla customizada
        self.custom_key_label = ttk.Label(
            button_frame,
            text=f"Tecla selecionada: {self.custom_key_name}",
            font=("Arial", 9),
            foreground="gray"
        )
        self.custom_key_label.pack(anchor=tk.W, padx=(20, 0), pady=(5, 10))
        
        # Botão para selecionar tecla customizada
        self.select_key_button = ttk.Button(
            button_frame,
            text="Selecionar Tecla Customizada",
            command=self._open_key_selector_dialog
        )
        self.select_key_button.pack(anchor=tk.W, padx=(20, 0))
        
        # ===== SEÇÃO: Coordenadas =====
        coord_frame = ttk.LabelFrame(main_frame, text="Definição de Local", padding="10")
        coord_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.coord_label = ttk.Label(
            coord_frame,
            text="Coordenadas: Não capturadas",
            font=("Arial", 10)
        )
        self.coord_label.pack(pady=(0, 10))
        
        self.capture_button = ttk.Button(
            coord_frame,
            text="Capturar Coordenada",
            command=self._start_capture_mode
        )
        self.capture_button.pack(fill=tk.X)
        
        # ===== SEÇÃO: Configuração de Delays =====
        timing_frame = ttk.LabelFrame(main_frame, text="Configuração de Tempo", padding="10")
        timing_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Delay entre cliques
        delay_frame = ttk.Frame(timing_frame)
        delay_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(delay_frame, text="Delay entre cliques (ms):", width=30).pack(side=tk.LEFT)
        delay_spinbox = ttk.Spinbox(
            delay_frame,
            from_=10,
            to=10000,
            textvariable=self.click_delay_ms,
            width=10,
            command=self._on_timing_change
        )
        delay_spinbox.pack(side=tk.LEFT, padx=10)
        
        # Duração de hold
        hold_frame = ttk.Frame(timing_frame)
        hold_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(hold_frame, text="Duração do hold (ms):", width=30).pack(side=tk.LEFT)
        hold_spinbox = ttk.Spinbox(
            hold_frame,
            from_=50,
            to=30000,
            textvariable=self.hold_duration_ms,
            width=10,
            command=self._on_timing_change
        )
        hold_spinbox.pack(side=tk.LEFT, padx=10)
        
        # ===== SEÇÃO: Hotkeys =====
        hotkey_frame = ttk.LabelFrame(main_frame, text="Hotkeys de Controle", padding="10")
        hotkey_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        hotkey_info = (
            f"{self._key_name(self.key_start)} - Iniciar (Start)\n"
            f"{self._key_name(self.key_pause)} - Pausar/Parar (Pause/Stop)\n"
            f"{self._key_name(self.key_exit)} - Sair (Exit)\n\n"
            "Dica: Configure a coordenada antes de iniciar!\n"
            "Acesse Configurações > Rebindar Teclas para mudar os hotkeys.\n"
            "Você pode minimizar a janela durante a execução."
        )
        
        self.hotkey_label = ttk.Label(
            hotkey_frame,
            text=hotkey_info,
            font=("Arial", 9),
            justify=tk.LEFT
        )
        self.hotkey_label.pack(anchor=tk.W, fill=tk.BOTH, expand=True)
        
        # ===== SEÇÃO: Status =====
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding="10")
        status_frame.pack(fill=tk.X)
        
        self.status_label = ttk.Label(
            status_frame,
            text="Status: Pronto",
            font=("Arial", 9)
        )
        self.status_label.pack(anchor=tk.W)
    
    def _on_action_change(self):
        """Callback quando tipo de ação é alterado."""
        self.config_mgr.set("action_type", self.action_type.get())
    
    def _on_timing_change(self):
        """Callback quando timing é alterado."""
        self.config_mgr.set("click_delay_ms", self.click_delay_ms.get())
        self.config_mgr.set("hold_duration_ms", self.hold_duration_ms.get())
    
    def _start_capture_mode(self):
        """Ativa o modo de captura de coordenadas."""
        messagebox.showinfo(
            "Captura de Coordenadas",
            "Clique em qualquer ponto da tela para capturar a coordenada.\n"
            "O capture mode será desativado automaticamente após o clique."
        )
        
        self.capture_mode = True
        self.capture_button.config(state=tk.DISABLED, text="Aguardando clique...")
        self._update_status("Aguardando clique na tela...", self.theme["warning"])
        self.root.update()
    
    def _on_key_press(self, key):
        """Manipulador de eventos de teclado para hotkeys."""
        try:
            # Normalizar tecla para comparação
            key_pressed = self._normalize_key(key)
            key_start = self._normalize_key(self.key_start)
            key_pause = self._normalize_key(self.key_pause)
            key_exit = self._normalize_key(self.key_exit)
            
            if key_pressed == key_start:
                if self.saved_x is None or self.saved_y is None:
                    messagebox.showwarning(
                        "Aviso",
                        "Por favor, capture uma coordenada primeiro!"
                    )
                    return
                
                self.is_running = True
                self.is_paused = False
                self._update_status("Executando...", self.theme["success"])
                
                self.macro_thread = threading.Thread(target=self._execute_macro, daemon=True)
                self.macro_thread.start()
            
            elif key_pressed == key_pause:
                self.is_running = False
                self.is_paused = True
                self._release_all()
                self._update_status("Pausado", self.theme["warning"])
            
            elif key_pressed == key_exit:
                self.is_running = False
                self.is_paused = False
                self._release_all()
                self.config_mgr.save_config()
                self.root.after(500, self._on_closing)
        
        except AttributeError:
            pass
    
    def _execute_macro(self):
        """Executa a automação do macro em thread separada."""
        try:
            action = self.action_type.get()
            button_type = self.button_type.get()
            
            # Mover mouse para coordenada
            if button_type in ["esquerdo", "direito"]:
                self.mouse_controller.position = (self.saved_x, self.saved_y)
                time.sleep(0.05)
            
            while self.is_running:
                if button_type == "esquerdo":
                    self._perform_action(Button.left, action)
                elif button_type == "direito":
                    self._perform_action(Button.right, action)
                elif button_type == "custom":
                    if self.custom_key:
                        self._perform_keyboard_action(self.custom_key, action)
                    else:
                        self._update_status("Erro: Nenhuma tecla customizada selecionada", self.theme["error"])
                        break
                
                # Delay entre cliques
                if action == "click":
                    time.sleep(self.click_delay_ms.get() / 1000)
                else:
                    # Se for hold, aguarda o tempo do hold + delay
                    time.sleep((self.hold_duration_ms.get() + self.click_delay_ms.get()) / 1000)
            
            self._release_all()
            self._update_status("Parado", self.theme["error"])
        
        except Exception as e:
            self._update_status(f"Erro: {str(e)}", self.theme["error"])
            print(f"Erro durante execução: {str(e)}")
    
    def _perform_action(self, button, action_type):
        """Realiza ação com o mouse."""
        try:
            if action_type == "click":
                self.mouse_controller.click(button, 1)
            else:  # hold
                self.mouse_controller.press(button)
                time.sleep(self.hold_duration_ms.get() / 1000)
                self.mouse_controller.release(button)
        except Exception as e:
            print(f"Erro ao performar ação do mouse: {e}")
    
    def _perform_keyboard_action(self, key, action_type):
        """Realiza ação com o teclado."""
        try:
            if action_type == "click":
                self.keyboard_controller.press(key)
                time.sleep(0.05)
                self.keyboard_controller.release(key)
            else:  # hold
                self.keyboard_controller.press(key)
                time.sleep(self.hold_duration_ms.get() / 1000)
                self.keyboard_controller.release(key)
        except Exception as e:
            print(f"Erro ao performar ação do teclado: {e}")
    
    def _release_all(self):
        """Solta todas as teclas e botões pressionados."""
        try:
            self.mouse_controller.release(Button.left)
        except:
            pass
        
        try:
            self.mouse_controller.release(Button.right)
        except:
            pass
        
        # Soltar tecla customizada
        try:
            if self.custom_key:
                self.keyboard_controller.release(self.custom_key)
        except:
            pass
    
    def _update_status(self, message, color):
        """Atualiza o label de status na GUI."""
        try:
            self.status_label.config(text=f"Status: {message}")
            self.root.update()
        except:
            pass
    
    def _on_mouse_click(self, x, y, button, pressed):
        """Manipulador de eventos de mouse para captura de coordenadas."""
        if self.capture_mode and pressed:
            self.saved_x = x
            self.saved_y = y
            self.capture_mode = False
            
            # Salvar coordenadas
            self.config_mgr.set("saved_x", self.saved_x)
            self.config_mgr.set("saved_y", self.saved_y)
            
            self.coord_label.config(
                text=f"Coordenadas: X={self.saved_x}, Y={self.saved_y}"
            )
            self.capture_button.config(state=tk.NORMAL, text="Capturar Coordenada")
            self._update_status("Coordenada capturada!", self.theme["success"])
    
    def _key_name(self, key):
        """Retorna o nome legível da tecla."""
        try:
            return key.name.upper()
        except:
            return str(key).upper()
    
    def _open_keybind_dialog(self):
        """Abre diálogo para rebindar teclas."""
        # Parar listener global para evitar conflitos
        try:
            if self.listener:
                self.listener.stop()
        except:
            pass
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Rebindar Teclas")
        dialog.geometry("450x350")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Aplicar tema ao diálogo
        dialog.configure(bg=self.theme["bg"])
        
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Pressione a tecla desejada para cada ação:", font=("Arial", 10)).pack(pady=(0, 20))
        
        # Tecla de Start
        start_frame = ttk.Frame(main_frame)
        start_frame.pack(fill=tk.X, pady=10)
        ttk.Label(start_frame, text="Iniciar (Start):", width=25).pack(side=tk.LEFT)
        start_button = ttk.Button(start_frame, text=self._key_name(self.key_start), width=15)
        start_button.pack(side=tk.LEFT, padx=10)
        start_button.config(command=lambda: self._capture_key("start", start_button))
        
        # Tecla de Pause
        pause_frame = ttk.Frame(main_frame)
        pause_frame.pack(fill=tk.X, pady=10)
        ttk.Label(pause_frame, text="Pausar/Parar (Pause):", width=25).pack(side=tk.LEFT)
        pause_button = ttk.Button(pause_frame, text=self._key_name(self.key_pause), width=15)
        pause_button.pack(side=tk.LEFT, padx=10)
        pause_button.config(command=lambda: self._capture_key("pause", pause_button))
        
        # Tecla de Exit
        exit_frame = ttk.Frame(main_frame)
        exit_frame.pack(fill=tk.X, pady=10)
        ttk.Label(exit_frame, text="Sair (Exit):", width=25).pack(side=tk.LEFT)
        exit_button = ttk.Button(exit_frame, text=self._key_name(self.key_exit), width=15)
        exit_button.pack(side=tk.LEFT, padx=10)
        exit_button.config(command=lambda: self._capture_key("exit", exit_button))
        
        # Botões
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        ttk.Button(button_frame, text="OK", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Restaurar Padrão", command=lambda: [self._reset_keybinds(), dialog.destroy()]).pack(side=tk.LEFT, padx=5)
        
        # Reiniciar listener quando dialog fecha
        def on_dialog_close():
            try:
                if not self.listener or not self.listener.is_alive():
                    self._initialize_listeners()
            except:
                self._initialize_listeners()
        
        dialog.protocol("WM_DELETE_WINDOW", lambda: [on_dialog_close(), dialog.destroy()])
    
    def _capture_key(self, action, button):
        """Captura uma tecla pressionada."""
        button.config(text="Aguardando tecla...", state=tk.DISABLED)
        self.root.update()
        
        def capture_in_thread():
            def on_press(key):
                try:
                    if action == "start":
                        self.key_start = key
                        key_str = self._key_to_string(key)
                        self.config_mgr.set("key_start", key_str)
                        print(f"Rebinded START para: {key_str}")
                    elif action == "pause":
                        self.key_pause = key
                        key_str = self._key_to_string(key)
                        self.config_mgr.set("key_pause", key_str)
                        print(f"Rebinded PAUSE para: {key_str}")
                    elif action == "exit":
                        self.key_exit = key
                        key_str = self._key_to_string(key)
                        self.config_mgr.set("key_exit", key_str)
                        print(f"Rebinded EXIT para: {key_str}")
                    
                    self.root.after(0, lambda: button.config(text=self._key_name(key), state=tk.NORMAL))
                    self.root.after(0, self._update_hotkey_display)
                    return False
                except:
                    return False
            
            try:
                with Listener(on_press=on_press) as listener:
                    listener.join(timeout=5)
            except:
                pass
            
            self.root.after(0, lambda: button.config(state=tk.NORMAL))
        
        capture_thread = threading.Thread(target=capture_in_thread, daemon=True)
        capture_thread.start()
    
    def _open_theme_dialog(self):
        """Abre diálogo para mudar o tema."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Escolher Tema")
        dialog.geometry("300x150")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        dialog.configure(bg=self.theme["bg"])
        
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Selecione um tema:", font=("Arial", 10)).pack(pady=(0, 20))
        
        def apply_theme(theme_name):
            self.current_theme = theme_name
            self.config_mgr.set("theme", theme_name)
            self.theme = ThemeManager.get_theme(theme_name)
            ThemeManager.configure_style(self.style, theme_name)
            self.root.configure(bg=self.theme["bg"])
            dialog.destroy()
            messagebox.showinfo("Tema Alterado", f"Tema '{theme_name}' aplicado com sucesso!")
        
        ttk.Button(main_frame, text="Tema Escuro", command=lambda: apply_theme("dark")).pack(fill=tk.X, pady=5)
        ttk.Button(main_frame, text="Tema Claro", command=lambda: apply_theme("light")).pack(fill=tk.X, pady=5)
    
    def _open_key_selector_dialog(self):
        """Abre diálogo para seleção de qualquer tecla do teclado."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Selecionar Tecla Customizada")
        dialog.geometry("450x250")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        dialog.configure(bg=self.theme["bg"])
        
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(
            main_frame,
            text="Pressione qualquer tecla do teclado\nque deseja usar no macro:",
            font=("Arial", 11),
            justify=tk.CENTER
        ).pack(pady=(0, 20))
        
        # Label para mostrar tecla pressionada
        key_display_label = ttk.Label(
            main_frame,
            text="Aguardando tecla...",
            font=("Arial", 14, "bold"),
            foreground="orange"
        )
        key_display_label.pack(pady=20)
        
        info_label = ttk.Label(
            main_frame,
            text="Exemplos: A, B, C, 1, 2, 3, Shift, Ctrl, Alt, etc.",
            font=("Arial", 9),
            foreground="gray"
        )
        info_label.pack(pady=10)
        
        def on_press(key):
            try:
                # Obter nome da tecla
                try:
                    key_name = key.name
                    if key_name.startswith("_"):
                        key_name = key_name[1:]  # Remove underscore
                except AttributeError:
                    # Para caracteres normais
                    key_name = str(key).replace("'", "")
                
                self.custom_key = key
                self.custom_key_name = key_name.upper()
                self.config_mgr.set("custom_key_name", self.custom_key_name)
                self.config_mgr.set("custom_key_stored", key_name.lower())
                
                # Atualizar display
                key_display_label.config(
                    text=f"Tecla selecionada: {self.custom_key_name}",
                    foreground="green"
                )
                
                # Atualizar label na janela principal
                self.custom_key_label.config(text=f"Tecla selecionada: {self.custom_key_name}")
                
                # Fechar diálogo após 1 segundo
                dialog.after(1000, dialog.destroy)
                
                return False  # Parar de capturar
            except Exception as e:
                print(f"Erro ao selecionar tecla: {e}")
                key_display_label.config(
                    text=f"Erro: {str(e)}",
                    foreground="red"
                )
                return False
        
        # Iniciar captura em thread separada
        def capture_in_thread():
            try:
                with Listener(on_press=on_press) as listener:
                    listener.join(timeout=15)  # Timeout de 15 segundos
            except:
                pass
            
            # Fechar diálogo se timeout
            try:
                dialog.destroy()
            except:
                pass
        
        capture_thread = threading.Thread(target=capture_in_thread, daemon=True)
        capture_thread.start()
    
    def _reset_keybinds(self):
        """Restaura as teclas padrão."""
        self.key_start = Key.f1
        self.key_pause = Key.f2
        self.key_exit = Key.f3
        self.config_mgr.set("key_start", "f1")
        self.config_mgr.set("key_pause", "f2")
        self.config_mgr.set("key_exit", "f3")
        self._update_hotkey_display()
    
    def _reset_all(self):
        """Restaura todas as configurações ao padrão."""
        if messagebox.askyesno("Confirmação", "Restaurar todas as configurações ao padrão?"):
            self._reset_keybinds()
            self.click_delay_ms.set(100)
            self.hold_duration_ms.set(500)
            self.action_type.set("click")
            self.button_type.set("esquerdo")
            self.custom_key = None
            self.custom_key_name = "Nenhuma"
            self.current_theme = "dark"
            self.config_mgr.set("theme", "dark")
            self.config_mgr.set("action_type", "click")
            self.config_mgr.set("click_delay_ms", 100)
            self.config_mgr.set("hold_duration_ms", 500)
            self.config_mgr.set("button_type", "esquerdo")
            self.config_mgr.set("custom_key", None)
            self.config_mgr.set("custom_key_name", "Nenhuma")
            self.theme = ThemeManager.get_theme("dark")
            ThemeManager.configure_style(self.style, "dark")
            self.root.configure(bg=self.theme["bg"])
            self.custom_key_label.config(text="Tecla selecionada: Nenhuma")
            messagebox.showinfo("Sucesso", "Configurações restauradas!")
    
    def _update_hotkey_display(self):
        """Atualiza a exibição dos hotkeys na GUI."""
        try:
            hotkey_info = (
                f"{self._key_name(self.key_start)} - Iniciar (Start)\n"
                f"{self._key_name(self.key_pause)} - Pausar/Parar (Pause/Stop)\n"
                f"{self._key_name(self.key_exit)} - Sair (Exit)\n\n"
                "Dica: Configure a coordenada antes de iniciar!\n"
                "Acesse Configurações > Rebindar Teclas para mudar os hotkeys.\n"
                "Você pode minimizar a janela durante a execução."
            )
            self.hotkey_label.config(text=hotkey_info)
            self.root.update()
        except:
            pass
    
    def _show_about(self):
        """Mostra a janela de sobre."""
        about_text = """Macro Automation v2.0

Automação profissional de mouse e teclado com GUI

Desenvolvido em Python com:
- tkinter (Interface Gráfica)
- pynput (Controle do Mouse e Teclado)
- JSON (Persistência de configurações)

Autor: Senior Python Developer
Data: 2026-01-06

Features V2.0:
✓ Interface moderna com tema claro/escuro
✓ Suporte a mouse e teclado
✓ Pressionar e manter (Hold) configurável
✓ Delay entre cliques em milissegundos
✓ Salvamento automático de configurações
✓ Hotkeys personalizáveis
✓ Execução em thread separada

Hotkeys Padrão:
F1 - Iniciar
F2 - Pausar
F3 - Sair
"""
        messagebox.showinfo("Sobre", about_text)
    
    def _save_config_manually(self):
        """Salva a configuração manualmente."""
        if self.config_mgr.save_config():
            messagebox.showinfo("Sucesso", f"Configuração salva em:\n{self.config_mgr.config_file}")
        else:
            messagebox.showerror("Erro", "Falha ao salvar configuração")
    
    def _create_default_config(self):
        """Cria um arquivo de configuração padrão que pode ser editado manualmente."""
        default_template = {
            "theme": "dark",
            "button_type": "esquerdo",
            "saved_x": None,
            "saved_y": None,
            "key_start": "f1",
            "key_pause": "f2",
            "key_exit": "f3",
            "click_delay_ms": 100,
            "action_type": "click",
            "hold_duration_ms": 500,
            "custom_key_name": "Nenhuma"
        }
        
        try:
            with open(self.config_mgr.config_file, 'w', encoding='utf-8') as f:
                json.dump(default_template, f, indent=4, ensure_ascii=False)
            messagebox.showinfo(
                "Sucesso",
                f"Arquivo de configuração padrão criado em:\n{self.config_mgr.config_file}\n\n"
                "Você pode editar este arquivo com um editor de texto se desejar."
            )
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao criar config padrão: {e}")
    
    def _on_closing(self):
        """Encerra a aplicação de forma segura."""
        self.is_running = False
        self._release_all()
        
        # Salvar todas as configurações atuais
        self.config_mgr.set("button_type", self.button_type.get())
        self.config_mgr.set("action_type", self.action_type.get())
        self.config_mgr.set("click_delay_ms", self.click_delay_ms.get())
        self.config_mgr.set("hold_duration_ms", self.hold_duration_ms.get())
        self.config_mgr.set("custom_key_name", self.custom_key_name)
        
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
        
        self.root.destroy()


def main():
    """Função principal para inicializar a aplicação."""
    try:
        # Setup logging para debug
        log_file = os.path.join(os.getcwd(), "macro_debug.log")
        with open(log_file, "a") as f:
            f.write(f"\n=== Iniciando em {os.getcwd()} ===\n")
        
        root = tk.Tk()
        app = MacroAutomation(root)
        root.mainloop()
    
    except Exception as e:
        print(f"Erro ao inicializar aplicação: {str(e)}")
        try:
            log_file = os.path.join(os.getcwd(), "macro_debug.log")
            with open(log_file, "a") as f:
                f.write(f"ERRO: {str(e)}\n")
        except:
            pass
        exit(1)


if __name__ == "__main__":
    main()
