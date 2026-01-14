"""
Script de Automação de Mouse e Teclado (Macro) - V2.0 PyQt6
Desenvolvido com PyQt6 para GUI moderna e pynput para controle
Autor: Senior Python Developer
Data: 2026-01-06

Interface profissional e moderna com PyQt6
"""

import sys
import threading
import time
import json
import os
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QRadioButton, QButtonGroup, QSpinBox,
    QGroupBox, QMessageBox, QDialog, QComboBox
)
from PyQt6.QtCore import Qt, pyqtSignal, QObject, QTimer
from PyQt6.QtGui import QFont, QIcon, QColor

try:
    from pynput.mouse import Controller as MouseController, Button, Listener as MouseListener
    from pynput.keyboard import Controller as KeyboardController, Listener, Key
except ImportError:
    print("Erro: pynput não está instalado. Execute: pip install pynput")
    exit(1)


class SignalEmitter(QObject):
    """Emite sinais para atualização segura da GUI da thread."""
    status_changed = pyqtSignal(str, str)
    coordinates_updated = pyqtSignal(int, int)


class ConfigManager:
    """Gerencia salvamento e carregamento de configurações."""
    
    def __init__(self, config_file="macro_config.json"):
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
            "action_type": "click",
            "hold_duration_ms": 500,
            "custom_key_name": "Nenhuma"
        }
        self.config = self.load_config()
    
    def load_config(self):
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    return {**self.default_config, **config}
            except Exception as e:
                print(f"Erro ao carregar configurações: {e}")
                return self.default_config.copy()
        return self.default_config.copy()
    
    def save_config(self):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Erro ao salvar configurações: {e}")
            return False
    
    def get(self, key, default=None):
        return self.config.get(key, default)
    
    def set(self, key, value):
        self.config[key] = value
        self.save_config()


class ThemeManager:
    """Gerencia temas da aplicação."""
    
    THEMES = {
        "dark": {
            "bg": "#1e1e1e",
            "fg": "#ffffff",
            "primary": "#0078d4",
            "success": "#4ec9b0",
            "warning": "#dcdcaa",
            "error": "#f48771",
            "widget_bg": "#2d2d2d",
            "border": "#3e3e3e"
        },
        "light": {
            "bg": "#ffffff",
            "fg": "#000000",
            "primary": "#0078d4",
            "success": "#008000",
            "warning": "#ff8c00",
            "error": "#ff0000",
            "widget_bg": "#f5f5f5",
            "border": "#cccccc"
        }
    }
    
    @classmethod
    def get_stylesheet(cls, theme_name):
        """Retorna stylesheet para o tema."""
        theme = cls.THEMES.get(theme_name, cls.THEMES["dark"])
        
        return f"""
            QMainWindow {{
                background-color: {theme['bg']};
                color: {theme['fg']};
            }}
            
            QWidget {{
                background-color: {theme['bg']};
                color: {theme['fg']};
            }}
            
            QGroupBox {{
                background-color: {theme['widget_bg']};
                border: 2px solid {theme['border']};
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }}
            
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }}
            
            QPushButton {{
                background-color: {theme['primary']};
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
                font-weight: bold;
            }}
            
            QPushButton:hover {{
                background-color: {theme['primary']};
                opacity: 0.8;
            }}
            
            QPushButton:pressed {{
                background-color: {theme['primary']};
                opacity: 0.6;
            }}
            
            QSpinBox {{
                background-color: {theme['widget_bg']};
                color: {theme['fg']};
                border: 1px solid {theme['border']};
                border-radius: 3px;
                padding: 5px;
            }}
            
            QComboBox {{
                background-color: {theme['widget_bg']};
                color: {theme['fg']};
                border: 1px solid {theme['border']};
                border-radius: 3px;
                padding: 5px;
            }}
            
            QLabel {{
                color: {theme['fg']};
            }}
            
            QRadioButton {{
                color: {theme['fg']};
            }}
            
            QMenuBar {{
                background-color: {theme['widget_bg']};
                color: {theme['fg']};
                border-bottom: 1px solid {theme['border']};
            }}
            
            QMenuBar::item:selected {{
                background-color: {theme['primary']};
            }}
            
            QMenu {{
                background-color: {theme['widget_bg']};
                color: {theme['fg']};
                border: 1px solid {theme['border']};
            }}
            
            QMenu::item:selected {{
                background-color: {theme['primary']};
            }}
        """
    
    @classmethod
    def get_theme(cls, theme_name):
        return cls.THEMES.get(theme_name, cls.THEMES["dark"])


class MacroAutomationPyQt(QMainWindow):
    """Aplicação principal com PyQt6."""
    
    def __init__(self):
        super().__init__()
        
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
        print(f"Arquivo de config: {self.config_mgr.config_file}")
        
        self.current_theme = self.config_mgr.get("theme", "dark")
        self.theme = ThemeManager.get_theme(self.current_theme)
        
        # Variáveis de controle
        self.mouse_controller = MouseController()
        self.keyboard_controller = KeyboardController()
        self.is_running = False
        self.is_paused = False
        self.capture_mode = False
        
        # Sinais para atualização segura da GUI
        self.signal_emitter = SignalEmitter()
        self.signal_emitter.status_changed.connect(self.update_status)
        self.signal_emitter.coordinates_updated.connect(self.update_coordinates_display)
        
        # Coordenadas
        self.saved_x = self.config_mgr.get("saved_x")
        self.saved_y = self.config_mgr.get("saved_y")
        
        # Configurações
        self.action_type = self.config_mgr.get("action_type", "click")
        self.button_type = self.config_mgr.get("button_type", "esquerdo")
        self.click_delay_ms = self.config_mgr.get("click_delay_ms", 100)
        self.hold_duration_ms = self.config_mgr.get("hold_duration_ms", 500)
        
        # Tecla customizada
        self.custom_key = self.config_mgr.get("custom_key", None)
        self.custom_key_name = self.config_mgr.get("custom_key_name", "Nenhuma")
        
        # Hotkeys
        key_start_str = self.config_mgr.get("key_start", "f1")
        key_pause_str = self.config_mgr.get("key_pause", "f2")
        key_exit_str = self.config_mgr.get("key_exit", "f3")
        print(f"Carregando hotkeys: START={key_start_str}, PAUSE={key_pause_str}, EXIT={key_exit_str}")
        
        self.key_start = self._string_to_key(key_start_str)
        self.key_pause = self._string_to_key(key_pause_str)
        self.key_exit = self._string_to_key(key_exit_str)
        
        print(f"Hotkeys convertidos: START={self.key_start}, PAUSE={self.key_pause}, EXIT={self.key_exit}")
        
        # Listeners
        self.listener = None
        self.mouse_listener = None
        
        self._initialize_listeners()
        self._init_ui()
        self.apply_theme(self.current_theme)
    
    def _string_to_key(self, key_str):
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
                        return key_str_clean
                    raise
        except Exception as e:
            print(f"_string_to_key erro: {key_str} -> retornando Key.f1 (erro: {e})")
            return Key.f1
    
    def _key_to_string(self, key):
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
        try:
            self.mouse_listener = MouseListener(on_click=self._on_mouse_click)
            self.mouse_listener.start()
            
            self.listener = Listener(on_press=self._on_key_press)
            self.listener.start()
        except Exception as e:
            print(f"Erro ao inicializar listeners: {e}")
    
    def _init_ui(self):
        """Inicializa a interface do usuário."""
        self.setWindowTitle("Macro Automation V2.0 - PyQt6")
        self.setGeometry(100, 100, 800, 700)
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout()
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Seção de tipo de ação
        action_group = QGroupBox("Tipo de Ação")
        action_layout = QVBoxLayout()
        
        self.action_button_group = QButtonGroup()
        action_click_radio = QRadioButton("Clique Único")
        action_hold_radio = QRadioButton("Pressionar e Manter (Hold)")
        
        if self.action_type == "click":
            action_click_radio.setChecked(True)
        else:
            action_hold_radio.setChecked(True)
        
        self.action_button_group.addButton(action_click_radio, 0)
        self.action_button_group.addButton(action_hold_radio, 1)
        self.action_button_group.buttonClicked.connect(self._on_action_changed)
        
        action_layout.addWidget(action_click_radio)
        action_layout.addWidget(action_hold_radio)
        action_group.setLayout(action_layout)
        main_layout.addWidget(action_group)
        
        # Seção de seleção de botão/tecla
        button_group = QGroupBox("Seleção de Botão/Tecla")
        button_layout = QVBoxLayout()
        
        self.button_button_group = QButtonGroup()
        left_button_radio = QRadioButton("Botão Esquerdo do Mouse")
        right_button_radio = QRadioButton("Botão Direito do Mouse")
        custom_button_radio = QRadioButton("Qualquer Tecla (Customizada)")
        
        button_map = {
            "esquerdo": 0,
            "direito": 1,
            "custom": 2
        }
        
        self.button_button_group.addButton(left_button_radio, 0)
        self.button_button_group.addButton(right_button_radio, 1)
        self.button_button_group.addButton(custom_button_radio, 2)
        
        if self.button_type in button_map:
            self.button_button_group.button(button_map[self.button_type]).setChecked(True)
        
        button_layout.addWidget(left_button_radio)
        button_layout.addWidget(right_button_radio)
        button_layout.addWidget(custom_button_radio)
        
        # Label para exibir tecla customizada
        self.custom_key_label = QLabel(f"Tecla selecionada: {self.custom_key_name}")
        self.custom_key_label.setStyleSheet("color: gray; font-size: 9pt;")
        button_layout.addWidget(self.custom_key_label)
        
        # Botão para selecionar tecla
        self.select_key_button = QPushButton("Selecionar Tecla Customizada")
        self.select_key_button.clicked.connect(self._open_key_selector_dialog)
        button_layout.addWidget(self.select_key_button)
        
        button_group.setLayout(button_layout)
        main_layout.addWidget(button_group)
        
        # Seção de coordenadas
        coord_group = QGroupBox("Definição de Local")
        coord_layout = QVBoxLayout()
        
        self.coord_label = QLabel("Coordenadas: Não capturadas")
        self.coord_label.setFont(QFont("Arial", 10))
        
        self.capture_button = QPushButton("Capturar Coordenada")
        self.capture_button.clicked.connect(self._start_capture_mode)
        
        if self.saved_x is not None and self.saved_y is not None:
            self.coord_label.setText(f"Coordenadas: X={self.saved_x}, Y={self.saved_y}")
        
        coord_layout.addWidget(self.coord_label)
        coord_layout.addWidget(self.capture_button)
        coord_group.setLayout(coord_layout)
        main_layout.addWidget(coord_group)
        
        # Seção de timing
        timing_group = QGroupBox("Configuração de Tempo")
        timing_layout = QVBoxLayout()
        
        # Delay entre cliques
        delay_layout = QHBoxLayout()
        delay_label = QLabel("Delay entre cliques (ms):")
        delay_label.setFixedWidth(200)
        self.delay_spinbox = QSpinBox()
        self.delay_spinbox.setMinimum(1)
        self.delay_spinbox.setMaximum(10000)
        self.delay_spinbox.setValue(self.click_delay_ms)
        self.delay_spinbox.valueChanged.connect(self._on_timing_changed)
        
        delay_layout.addWidget(delay_label)
        delay_layout.addWidget(self.delay_spinbox)
        delay_layout.addStretch()
        timing_layout.addLayout(delay_layout)
        
        # Hold duration
        hold_layout = QHBoxLayout()
        hold_label = QLabel("Duração do hold (ms):")
        hold_label.setFixedWidth(200)
        self.hold_spinbox = QSpinBox()
        self.hold_spinbox.setMinimum(50)
        self.hold_spinbox.setMaximum(999999999)
        self.hold_spinbox.setValue(self.hold_duration_ms)
        self.hold_spinbox.valueChanged.connect(self._on_timing_changed)
        
        hold_layout.addWidget(hold_label)
        hold_layout.addWidget(self.hold_spinbox)
        hold_layout.addStretch()
        timing_layout.addLayout(hold_layout)
        
        timing_group.setLayout(timing_layout)
        main_layout.addWidget(timing_group)
        
        # Seção de informações
        info_group = QGroupBox("Hotkeys de Controle")
        info_layout = QVBoxLayout()
        
        info_text = (
            f"{self._key_name(self.key_start)} - Iniciar (Start)\n"
            f"{self._key_name(self.key_pause)} - Pausar/Parar (Pause/Stop)\n"
            f"{self._key_name(self.key_exit)} - Sair (Exit)\n\n"
            "Configure a coordenada antes de iniciar!\n"
            "Acesse Configurações para rebindar as teclas."
        )
        
        self.info_label = QLabel(info_text)
        self.info_label.setFont(QFont("Arial", 9))
        info_layout.addWidget(self.info_label)
        info_group.setLayout(info_layout)
        main_layout.addWidget(info_group, 1)
        
        # Status
        status_group = QGroupBox("Status")
        status_layout = QVBoxLayout()
        self.status_label = QLabel("Status: Pronto")
        self.status_label.setFont(QFont("Arial", 10))
        status_layout.addWidget(self.status_label)
        status_group.setLayout(status_layout)
        main_layout.addWidget(status_group)
        
        central_widget.setLayout(main_layout)
        
        # Criar menu bar
        menubar = self.menuBar()
        
        # Menu Arquivo
        arquivo_menu = menubar.addMenu("Arquivo")
        save_config_action = arquivo_menu.addAction("Salvar Config Manualmente")
        save_config_action.triggered.connect(self._save_config_manually)
        
        create_config_action = arquivo_menu.addAction("Criar Config Padrão")
        create_config_action.triggered.connect(self._create_default_config)
        
        arquivo_menu.addSeparator()
        sair_action = arquivo_menu.addAction("Sair")
        sair_action.triggered.connect(self.close)
        
        # Menu Configurações
        config_menu = menubar.addMenu("Configurações")
        rebind_action = config_menu.addAction("Rebindar Teclas...")
        rebind_action.triggered.connect(self._open_keybind_dialog)
        
        theme_action = config_menu.addAction("Alterar Tema...")
        theme_action.triggered.connect(self._open_theme_dialog)
        
        config_menu.addSeparator()
        reset_action = config_menu.addAction("Restaurar Padrão")
        reset_action.triggered.connect(self._reset_all)
        
        # Menu Ajuda
        ajuda_menu = menubar.addMenu("Ajuda")
        sobre_action = ajuda_menu.addAction("Sobre")
        sobre_action.triggered.connect(self._show_about)
    
    def _start_capture_mode(self):
        """Ativa o modo de captura de coordenadas."""
        QMessageBox.information(
            self,
            "Captura de Coordenadas",
            "Clique em qualquer ponto da tela para capturar a coordenada.\n"
            "O capture mode será desativado automaticamente após o clique."
        )
        
        self.capture_mode = True
        self.capture_button.setEnabled(False)
        self.capture_button.setText("Aguardando clique...")
        self.signal_emitter.status_changed.emit("Aguardando clique na tela...", "warning")
    
    def _on_action_changed(self):
        """Callback quando tipo de ação muda."""
        if self.action_button_group.checkedId() == 0:
            self.action_type = "click"
        else:
            self.action_type = "hold"
        self.config_mgr.set("action_type", self.action_type)
    
    def _on_timing_changed(self):
        """Callback quando timing muda."""
        self.click_delay_ms = self.delay_spinbox.value()
        self.hold_duration_ms = self.hold_spinbox.value()
        self.config_mgr.set("click_delay_ms", self.click_delay_ms)
        self.config_mgr.set("hold_duration_ms", self.hold_duration_ms)
    
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
                    QMessageBox.warning(
                        self,
                        "Aviso",
                        "Por favor, capture uma coordenada primeiro!"
                    )
                    return
                
                self.is_running = True
                self.is_paused = False
                self.signal_emitter.status_changed.emit("Executando...", "success")
                
                thread = threading.Thread(target=self._execute_macro, daemon=True)
                thread.start()
            
            elif key_pressed == key_pause:
                self.is_running = False
                self.is_paused = True
                self._release_all()
                self.signal_emitter.status_changed.emit("Pausado", "warning")
            
            elif key_pressed == key_exit:
                self.is_running = False
                self.is_paused = False
                self._release_all()
                self.config_mgr.save_config()
                QTimer.singleShot(500, self.close)
        
        except AttributeError:
            pass
    
    def _execute_macro(self):
        """Executa a automação do macro."""
        try:
            # Atualizar button_type a partir do selecionado
            button_map = {0: "esquerdo", 1: "direito", 2: "custom"}
            self.button_type = button_map.get(self.button_button_group.checkedId(), "esquerdo")
            self.config_mgr.set("button_type", self.button_type)
            
            while self.is_running:
                if self.button_type in ["esquerdo", "direito"]:
                    self.mouse_controller.position = (self.saved_x, self.saved_y)
                    time.sleep(0.05)
                
                if self.button_type == "esquerdo":
                    self._perform_action(Button.left, self.action_type)
                elif self.button_type == "direito":
                    self._perform_action(Button.right, self.action_type)
                elif self.button_type == "custom":
                    if self.custom_key:
                        self._perform_keyboard_action(self.custom_key, self.action_type)
                    else:
                        self.signal_emitter.status_changed.emit("Erro: Nenhuma tecla customizada selecionada", "error")
                        break
                
                if self.action_type == "click":
                    time.sleep(self.click_delay_ms / 1000)
                else:
                    time.sleep((self.hold_duration_ms + self.click_delay_ms) / 1000)
            
            self._release_all()
            self.signal_emitter.status_changed.emit("Parado", "error")
        
        except Exception as e:
            self.signal_emitter.status_changed.emit(f"Erro: {str(e)}", "error")
            print(f"Erro durante execução: {str(e)}")
    
    def _perform_action(self, button, action_type):
        try:
            if action_type == "click":
                self.mouse_controller.click(button, 1)
            else:
                self.mouse_controller.press(button)
                time.sleep(self.hold_duration_ms / 1000)
                self.mouse_controller.release(button)
        except Exception as e:
            print(f"Erro ao performar ação do mouse: {e}")
    
    def _perform_keyboard_action(self, key, action_type):
        try:
            if action_type == "click":
                self.keyboard_controller.press(key)
                time.sleep(0.05)
                self.keyboard_controller.release(key)
            else:
                self.keyboard_controller.press(key)
                time.sleep(self.hold_duration_ms / 1000)
                self.keyboard_controller.release(key)
        except Exception as e:
            print(f"Erro ao performar ação do teclado: {e}")
    
    def _release_all(self):
        try:
            self.mouse_controller.release(Button.left)
        except:
            pass
        
        try:
            self.mouse_controller.release(Button.right)
        except:
            pass
        
        if self.custom_key:
            try:
                self.keyboard_controller.release(self.custom_key)
            except:
                pass
    
    def _on_mouse_click(self, x, y, button, pressed):
        if self.capture_mode and pressed:
            self.saved_x = x
            self.saved_y = y
            self.capture_mode = False
            
            self.config_mgr.set("saved_x", self.saved_x)
            self.config_mgr.set("saved_y", self.saved_y)
            
            self.signal_emitter.coordinates_updated.emit(self.saved_x, self.saved_y)
            self.capture_button.setEnabled(True)
            self.capture_button.setText("Capturar Coordenada")
            self.signal_emitter.status_changed.emit("Coordenada capturada!", "success")
    
    def update_coordinates_display(self, x, y):
        self.coord_label.setText(f"Coordenadas: X={x}, Y={y}")
    
    def update_status(self, message, color_type):
        self.status_label.setText(f"Status: {message}")
    
    def _key_name(self, key):
        try:
            return key.name.upper()
        except:
            return str(key).upper()
    
    def _open_keybind_dialog(self):
        """Abre diálogo para rebindar as teclas de hotkey (F1, F2, F3)."""
        # Parar listener global para evitar conflitos
        try:
            if self.listener:
                self.listener.stop()
        except:
            pass
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Rebindar Hotkeys")
        dialog.setGeometry(200, 200, 450, 300)
        
        layout = QVBoxLayout()
        
        title = QLabel("Selecione as teclas para os hotkeys:")
        title_font = title.font()
        title_font.setBold(True)
        title_font.setPointSize(11)
        title.setFont(title_font)
        layout.addWidget(title)
        
        # Start key
        start_layout = QHBoxLayout()
        start_label = QLabel("Iniciar (F1):")
        start_label.setMinimumWidth(100)
        self.start_key_display = QLabel(self._key_name(self.key_start))
        start_key_button = QPushButton("Selecionar")
        start_key_button.clicked.connect(lambda: self._capture_hotkey_rebind("start", dialog))
        start_layout.addWidget(start_label)
        start_layout.addWidget(self.start_key_display)
        start_layout.addWidget(start_key_button)
        layout.addLayout(start_layout)
        
        # Pause key
        pause_layout = QHBoxLayout()
        pause_label = QLabel("Pausar (F2):")
        pause_label.setMinimumWidth(100)
        self.pause_key_display = QLabel(self._key_name(self.key_pause))
        pause_key_button = QPushButton("Selecionar")
        pause_key_button.clicked.connect(lambda: self._capture_hotkey_rebind("pause", dialog))
        pause_layout.addWidget(pause_label)
        pause_layout.addWidget(self.pause_key_display)
        pause_layout.addWidget(pause_key_button)
        layout.addLayout(pause_layout)
        
        # Exit key
        exit_layout = QHBoxLayout()
        exit_label = QLabel("Sair (F3):")
        exit_label.setMinimumWidth(100)
        self.exit_key_display = QLabel(self._key_name(self.key_exit))
        exit_key_button = QPushButton("Selecionar")
        exit_key_button.clicked.connect(lambda: self._capture_hotkey_rebind("exit", dialog))
        exit_layout.addWidget(exit_label)
        exit_layout.addWidget(self.exit_key_display)
        exit_layout.addWidget(exit_key_button)
        layout.addLayout(exit_layout)
        
        layout.addStretch()
        
        close_button = QPushButton("Fechar")
        close_button.clicked.connect(dialog.accept)
        layout.addWidget(close_button)
        
        dialog.setLayout(layout)
        dialog.exec()
        
        # Reiniciar listener quando dialog fecha
        try:
            if not self.listener or not self.listener.is_alive():
                self._initialize_listeners()
        except:
            self._initialize_listeners()
    
    def _capture_hotkey_rebind(self, hotkey_type, parent_dialog):
        """Captura uma tecla para rebindar um hotkey."""
        dialog = QDialog(parent_dialog)
        dialog.setWindowTitle(f"Selecionar Tecla para {hotkey_type.capitalize()}")
        dialog.setGeometry(250, 250, 400, 200)
        dialog.setModal(True)
        
        layout = QVBoxLayout()
        
        title = QLabel(f"Pressione a tecla para {hotkey_type}...")
        title_font = title.font()
        title_font.setPointSize(12)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        
        key_display = QLabel("Aguardando...")
        key_display.setAlignment(Qt.AlignmentFlag.AlignCenter)
        key_display_font = key_display.font()
        key_display_font.setPointSize(14)
        key_display_font.setBold(True)
        key_display.setFont(key_display_font)
        layout.addWidget(key_display)
        
        dialog.setLayout(layout)
        
        listener_obj = None
        
        def capture_key_in_thread():
            nonlocal listener_obj
            
            def on_press(key):
                try:
                    # Obter nome da tecla
                    try:
                        key_name = key.name
                        if key_name.startswith("_"):
                            key_name = key_name[1:]
                    except AttributeError:
                        key_name = str(key).replace("'", "")
                    
                    key_name_clean = key_name.lower()
                    
                    # Salvar a tecla apropriada
                    if hotkey_type == "start":
                        self.key_start = key
                        self.config_mgr.set("key_start", key_name_clean)
                        self.start_key_display.setText(self._key_name(self.key_start))
                        print(f"Rebinded START para: {key_name_clean}")
                    elif hotkey_type == "pause":
                        self.key_pause = key
                        self.config_mgr.set("key_pause", key_name_clean)
                        self.pause_key_display.setText(self._key_name(self.key_pause))
                        print(f"Rebinded PAUSE para: {key_name_clean}")
                    elif hotkey_type == "exit":
                        self.key_exit = key
                        self.config_mgr.set("key_exit", key_name_clean)
                        self.exit_key_display.setText(self._key_name(self.key_exit))
                        print(f"Rebinded EXIT para: {key_name_clean}")
                    
                    # Atualizar display
                    key_display.setText(f"Tecla selecionada: {key_name.upper()}")
                    
                    # Fechar após 1 segundo
                    QTimer.singleShot(1000, dialog.accept)
                    return False
                except Exception as e:
                    print(f"Erro ao capturar hotkey: {e}")
                    return False
            
            try:
                listener_obj = Listener(on_press=on_press)
                listener_obj.start()
                listener_obj.join(timeout=15)
            except Exception as e:
                print(f"Erro na listener: {e}")
        
        thread = threading.Thread(target=capture_key_in_thread, daemon=True)
        thread.start()
        
        dialog.exec()
        
        if listener_obj:
            try:
                listener_obj.stop()
            except:
                pass
        
        # Atualizar info_label na janela principal
        info_text = (
            f"{self._key_name(self.key_start)} - Iniciar (Start)\n"
            f"{self._key_name(self.key_pause)} - Pausar/Parar (Pause/Stop)\n"
            f"{self._key_name(self.key_exit)} - Sair (Exit)\n\n"
            "Configure a coordenada antes de iniciar!\n"
            "Acesse Configurações para rebindar as teclas."
        )
        self.info_label.setText(info_text)
    
    def _open_theme_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Escolher Tema")
        dialog.setGeometry(200, 200, 300, 150)
        
        layout = QVBoxLayout()
        
        label = QLabel("Selecione um tema:")
        layout.addWidget(label)
        
        dark_button = QPushButton("Tema Escuro")
        dark_button.clicked.connect(lambda: self._apply_theme_dialog("dark", dialog))
        layout.addWidget(dark_button)
        
        light_button = QPushButton("Tema Claro")
        light_button.clicked.connect(lambda: self._apply_theme_dialog("light", dialog))
        layout.addWidget(light_button)
        
        dialog.setLayout(layout)
        dialog.exec()
    
    def _apply_theme_dialog(self, theme_name, dialog):
        self.apply_theme(theme_name)
        self.current_theme = theme_name
        self.config_mgr.set("theme", theme_name)
        dialog.close()
        QMessageBox.information(self, "Tema Alterado", f"Tema '{theme_name}' aplicado com sucesso!")
    
    def apply_theme(self, theme_name):
        stylesheet = ThemeManager.get_stylesheet(theme_name)
        self.setStyleSheet(stylesheet)
    
    def _reset_all(self):
        reply = QMessageBox.question(
            self,
            "Confirmação",
            "Restaurar todas as configurações ao padrão?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self.key_start = Key.f1
            self.key_pause = Key.f2
            self.key_exit = Key.f3
            self.config_mgr.set("key_start", "f1")
            self.config_mgr.set("key_pause", "f2")
            self.config_mgr.set("key_exit", "f3")
            
            self.custom_key = None
            self.custom_key_name = "Nenhuma"
            self.config_mgr.set("custom_key_name", "Nenhuma")
            self.custom_key_label.setText("Tecla selecionada: Nenhuma")
            
            self.delay_spinbox.setValue(100)
            self.hold_spinbox.setValue(500)
            self.action_button_group.button(0).setChecked(True)
            self.button_button_group.button(0).setChecked(True)
            
            self.config_mgr.set("theme", "dark")
            self.apply_theme("dark")
            
            QMessageBox.information(self, "Sucesso", "Configurações restauradas!")
    
    def _show_about(self):
        about_text = """Macro Automation v2.0 - PyQt6

Automação profissional de mouse e teclado com interface moderna

Desenvolvido em Python com:
- PyQt6 (Interface Gráfica Moderna)
- pynput (Controle do Mouse e Teclado)
- JSON (Persistência de configurações)

Autor: Senior Python Developer
Data: 2026-01-06

Features V2.0:
✓ Interface moderna com PyQt6
✓ Tema claro/escuro
✓ Suporte a mouse e teclado
✓ Press and Hold configurável
✓ Delay em milissegundos
✓ Salvamento automático de configurações
✓ Hotkeys personalizáveis
✓ Execução em thread separada

Hotkeys Padrão:
F1 - Iniciar
F2 - Pausar
F3 - Sair
"""
        QMessageBox.information(self, "Sobre", about_text)
    
    def _save_config_manually(self):
        """Salva a configuração manualmente."""
        if self.config_mgr.save_config():
            QMessageBox.information(self, "Sucesso", f"Configuração salva em:\n{self.config_mgr.config_file}")
        else:
            QMessageBox.critical(self, "Erro", "Falha ao salvar configuração")
    
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
            QMessageBox.information(
                self,
                "Sucesso",
                f"Arquivo de configuração padrão criado em:\n{self.config_mgr.config_file}\n\n"
                "Você pode editar este arquivo com um editor de texto se desejar."
            )
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao criar config padrão: {e}")
    
    def _open_key_selector_dialog(self):
        """Abre diálogo para seleção de qualquer tecla do teclado."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Selecionar Tecla Customizada")
        dialog.setGeometry(100, 100, 450, 250)
        
        layout = QVBoxLayout()
        
        title_label = QLabel("Pressione qualquer tecla do teclado\nque deseja usar no macro:")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font = title_label.font()
        font.setPointSize(11)
        title_label.setFont(font)
        layout.addWidget(title_label)
        
        selected_key_label = QLabel("Aguardando tecla...")
        selected_key_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font = selected_key_label.font()
        font.setPointSize(14)
        font.setBold(True)
        selected_key_label.setFont(font)
        layout.addWidget(selected_key_label)
        
        info_label = QLabel("Exemplos: A, B, C, 1, 2, 3, Shift, Ctrl, Alt, etc.")
        info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font = info_label.font()
        font.setPointSize(9)
        info_label.setFont(font)
        layout.addWidget(info_label)
        
        dialog.setLayout(layout)
        
        listener_obj = None
        
        def capture_key_in_thread():
            """Captura a tecla em uma thread separada."""
            nonlocal listener_obj
            
            def on_press(key):
                try:
                    # Obter nome da tecla de forma robusta
                    try:
                        key_name = key.name
                        if key_name.startswith("_"):
                            key_name = key_name[1:]
                    except AttributeError:
                        # Para caracteres normais/KeyCode
                        key_name = str(key).replace("'", "")
                    
                    self.custom_key = key
                    self.custom_key_name = key_name.upper()
                    self.config_mgr.set("custom_key_name", self.custom_key_name)
                    
                    # Atualizar UI
                    selected_key_label.setText(f"Tecla selecionada: {self.custom_key_name}")
                    self.custom_key_label.setText(f"Tecla selecionada: {self.custom_key_name}")
                    
                    # Fechar após 1 segundo
                    QTimer.singleShot(1000, dialog.accept)
                    return False
                except Exception as e:
                    print(f"Erro ao capturar tecla: {e}")
                    return False
            
            try:
                listener_obj = Listener(on_press=on_press)
                listener_obj.start()
                listener_obj.join(timeout=15)  # Timeout de 15 segundos
            except Exception as e:
                print(f"Erro na listener: {e}")
        
        thread = threading.Thread(target=capture_key_in_thread, daemon=True)
        thread.start()
        
        dialog.exec()
        
        if listener_obj:
            try:
                listener_obj.stop()
            except:
                pass
    
    def closeEvent(self, event):
        self.is_running = False
        self._release_all()
        
        self.config_mgr.set("button_type", "esquerdo")  # Salvar estado
        
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
        
        event.accept()


def main():
    app = QApplication(sys.argv)
    window = MacroAutomationPyQt()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
