import sys
import os
import time
import requests
import pymem
import pymem.process
import win32gui
import win32con
import win32api
import win32com.client

from PySide6 import QtWidgets, QtCore, QtGui
from PySide6.QtSvgWidgets import QSvgWidget, QGraphicsSvgItem
from PySide6.QtWidgets import (
    QDialog, QCheckBox, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QFrame, QWidget
)

from PySide6.QtMultimedia import QMediaPlayer, QAudioOutput
from PySide6.QtCore import QUrl

try:
    from pypresence import Presence
    HAS_RPC = True
except ImportError:
    HAS_RPC = False

# Función crucial para PyInstaller: Encuentra la ruta de los recursos extraídos
def get_resource_path(relative_path):
    try:
        # PyInstaller crea una carpeta temporal y guarda la ruta en _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# ----------------------------------------------------------------------
# OFFSETS & ENTITY LOGIC
# ----------------------------------------------------------------------
class Offsets:
    dwEntityList = None
    dwLocalPlayerController = None
    dwLocalPlayerPawn = None
    dwViewMatrix = None
    
    m_iszPlayerName = None
    m_iHealth = None
    m_ArmorValue = None
    m_iTeamNum = None
    m_lifeState = None
    m_vOldOrigin = None
    m_hPlayerPawn = None

    @classmethod
    def load(cls):
        try:
            offsets = requests.get("https://raw.githubusercontent.com/a2x/cs2-dumper/main/output/offsets.json").json()
            client_dll = requests.get("https://raw.githubusercontent.com/a2x/cs2-dumper/main/output/client_dll.json").json()
            
            cls.dwEntityList = offsets["client.dll"]["dwEntityList"]
            cls.dwLocalPlayerController = offsets["client.dll"]["dwLocalPlayerController"]
            cls.dwLocalPlayerPawn = offsets["client.dll"]["dwLocalPlayerPawn"]
            cls.dwViewMatrix = offsets["client.dll"]["dwViewMatrix"]
            
            classes = client_dll["client.dll"]["classes"]

            cls.m_iszPlayerName = classes["CBasePlayerController"]["fields"]["m_iszPlayerName"]
            cls.m_iHealth = classes["C_BaseEntity"]["fields"]["m_iHealth"]
            cls.m_ArmorValue = classes["C_CSPlayerPawn"]["fields"]["m_ArmorValue"]
            cls.m_iTeamNum = classes["C_BaseEntity"]["fields"]["m_iTeamNum"]
            cls.m_lifeState = classes["C_BaseEntity"]["fields"]["m_lifeState"]
            cls.m_vOldOrigin = classes["C_BasePlayerPawn"]["fields"]["m_vOldOrigin"]
            
            if "m_hPlayerPawn" in classes.get("CBasePlayerController", {}).get("fields", {}):
                cls.m_hPlayerPawn = classes["CBasePlayerController"]["fields"]["m_hPlayerPawn"]
            elif "m_hPawn" in classes.get("CBasePlayerController", {}).get("fields", {}):
                cls.m_hPlayerPawn = classes["CBasePlayerController"]["fields"]["m_hPawn"]
            else:
                cls.m_hPlayerPawn = 0x80C 
        except Exception as e:
            print(f"[Offsets] Error al cargar: {e}")

class Entity:
    def __init__(self, controller, pawn):
        self.controller = controller
        self.pawn = pawn
        self.health = 0
        self.team = 0
        self.lifestate = 0
        self.name = ""
        self.pos = None

class ESP:
    def __init__(self):
        self.pm = None
        self.client = None
        self.entities = []
        self.local_player = None

    def initialize(self) -> bool:
        try:
            self.pm = pymem.Pymem("cs2.exe")
            self.client = pymem.process.module_from_name(self.pm.process_handle, "client.dll").lpBaseOfDll
            Offsets.load()
            return True
        except Exception as e:
            print(f"[ESP Init] Error: {e}")
            return False

    def update_entities(self):
        self.entities.clear()
        try:
            local_controller = self.pm.read_ulonglong(self.client + Offsets.dwLocalPlayerController)
            local_pawn = self.pm.read_ulonglong(self.client + Offsets.dwLocalPlayerPawn)
            if not local_controller or not local_pawn: return
                
            self.local_player = Entity(local_controller, local_pawn)
            self.local_player.team = self.pm.read_int(local_pawn + Offsets.m_iTeamNum)
            self.local_player.pos = tuple(self.pm.read_float(local_pawn + Offsets.m_vOldOrigin + i * 4) for i in range(3))
            
            entity_list = self.pm.read_ulonglong(self.client + Offsets.dwEntityList)
            if not entity_list: return
                
            for i in range(1, 65):
                self._process_entity(entity_list, i, local_controller)
        except Exception:
            pass

    def _process_entity(self, entity_list, index, local_controller):
        try:
            list_entry = self.pm.read_ulonglong(entity_list + (8 * (index & 0x7FFF) >> 9) + 16)
            if not list_entry: return
                
            controller = self.pm.read_ulonglong(list_entry + 112 * (index & 0x1FF))
            if controller == local_controller: return
                
            pawn_handle = self.pm.read_ulonglong(controller + Offsets.m_hPlayerPawn)
            pawn_entry = self.pm.read_ulonglong(entity_list + (8 * ((pawn_handle & 0x7FFF) >> 9)) + 16)
            pawn = self.pm.read_ulonglong(pawn_entry + 112 * (pawn_handle & 0x1FF))
            if not pawn: return
                
            entity = Entity(controller, pawn)
            entity.health = self.pm.read_int(pawn + Offsets.m_iHealth)
            if not 0 < entity.health <= 100: return
                
            entity.team = self.pm.read_int(pawn + Offsets.m_iTeamNum)
            entity.lifestate = self.pm.read_int(pawn + Offsets.m_lifeState)
            entity.name = self.pm.read_string(controller + Offsets.m_iszPlayerName)
            entity.pos = tuple(self.pm.read_float(pawn + Offsets.m_vOldOrigin + i * 4) for i in range(3))
            
            self.entities.append(entity)
        except Exception:
            pass

# ----------------------------------------------------------------------
# GUI UI MENU PIGWARE 
# ----------------------------------------------------------------------
def get_stylesheet():
    return """
    QFrame#MainFrame {
        background-color: #121212;
        border: 1px solid #ff1d65;
        border-radius: 12px;
    }
    QFrame#HeaderFrame {
        background-color: #181818;
        border-top-left-radius: 11px;
        border-top-right-radius: 11px;
        border-bottom: 1px solid #282828;
    }
    QLabel#SectionTitle {
        color: #ff1d65;
        font-family: 'Segoe UI', sans-serif;
        font-size: 11px;
        font-weight: bold;
    }
    QCheckBox {
        color: #C0C0C0;
        font-family: 'Segoe UI', sans-serif;
        font-size: 11px;
        spacing: 10px;
    }
    QCheckBox::indicator {
        width: 16px;
        height: 16px;
        border-radius: 4px;
    }
    QCheckBox::indicator:unchecked {
        background-color: #1a1a1a;
        border: 1px solid #333333;
    }
    QCheckBox::indicator:checked {
        background-color: #ff1d65;
        border: 1px solid #ff1d65;
        image: none; 
    }
    QPushButton#UnloadBtn {
        background-color: #2a2a2a;
        color: #ff1d65;
        border: 1px solid #ff1d65;
        border-radius: 6px;
        font-family: 'Segoe UI', sans-serif;
        font-size: 11px;
        font-weight: bold;
        padding: 6px;
    }
    QPushButton#UnloadBtn:hover {
        background-color: #ff1d65;
        color: #FFFFFF;
    }
    """

class DraggableHeader(QFrame):
    def __init__(self, parent_dialog):
        super().__init__()
        self.dialog = parent_dialog
        self.drag_position = QtCore.QPoint()

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.drag_position = event.globalPosition().toPoint() - self.dialog.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == QtCore.Qt.LeftButton:
            self.dialog.move(event.globalPosition().toPoint() - self.drag_position)
            event.accept()

class OverlayMenu(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.Tool | QtCore.Qt.WindowStaysOnTopHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground) 
        self.resize(220, 230)
        self.setStyleSheet(get_stylesheet())
        self._setup_ui()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        self.main_frame = QFrame(self)
        self.main_frame.setObjectName("MainFrame")
        main_layout.addWidget(self.main_frame)

        frame_layout = QVBoxLayout(self.main_frame)
        frame_layout.setContentsMargins(0, 0, 0, 0)
        frame_layout.setSpacing(0)

        # Header (Logo SVG, Texto y Drag)
        header_frame = DraggableHeader(self)
        header_frame.setObjectName("HeaderFrame")
        header_frame.setFixedHeight(40)
        header_layout = QHBoxLayout(header_frame)
        header_layout.setContentsMargins(15, 0, 15, 0)
        header_layout.setSpacing(8) 

        logo_path = get_resource_path("logo.svg")
        self.logo = QSvgWidget(logo_path)
        self.logo.setFixedSize(20, 20)
        header_layout.addWidget(self.logo)

        lbl_header_title = QLabel("PIGWARE")
        lbl_header_title.setStyleSheet("color: #ff1d65; font-family: 'Segoe UI'; font-size: 12px; font-weight: bold;")
        header_layout.addWidget(lbl_header_title)

        header_layout.addStretch()
        frame_layout.addWidget(header_frame)

        # Contenido (Visuals)
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(20, 15, 20, 15)
        content_layout.setSpacing(12)

        lbl_esp_title = QLabel("VISUALS")
        lbl_esp_title.setObjectName("SectionTitle")
        content_layout.addWidget(lbl_esp_title)

        cb_box = QCheckBox("Box")
        cb_box.setChecked(True)
        cb_box.stateChanged.connect(lambda state: self._toggle_vis("draw_box", state))
        content_layout.addWidget(cb_box)

        cb_health = QCheckBox("Health")
        cb_health.setChecked(True)
        cb_health.stateChanged.connect(lambda state: self._toggle_vis("draw_health", state))
        content_layout.addWidget(cb_health)

        cb_name = QCheckBox("Name")
        cb_name.setChecked(True)
        cb_name.stateChanged.connect(lambda state: self._toggle_vis("draw_names", state))
        content_layout.addWidget(cb_name)

        content_layout.addStretch()

        btn_unload = QPushButton("UNLOAD")
        btn_unload.setObjectName("UnloadBtn")
        btn_unload.clicked.connect(self._unload)
        content_layout.addWidget(btn_unload)

        frame_layout.addWidget(content_widget)

    def _toggle_vis(self, key, state):
        # Reproducir sonido de click
        if hasattr(self.parent(), "click_player"):
            click_path = get_resource_path("click.mp3")
            self.parent().click_player.setSource(QUrl.fromLocalFile(click_path))
            self.parent().click_player.play()

        # Cambiar el valor visual
        if hasattr(self.parent(), "renderer"):
            setattr(self.parent().renderer, key, bool(state))

    def _unload(self):
        # Cierre y Unload Inmediato
        self.hide()
        if hasattr(self.parent(), "hide"):
            self.parent().hide()
        QtWidgets.QApplication.quit()

    def keyPressEvent(self, event):
        # Ignoramos la tecla ESC para que no cierre el menú
        if event.key() == QtCore.Qt.Key_Escape:
            event.ignore()
        else:
            super().keyPressEvent(event)

# ----------------------------------------------------------------------
# RENDERER & OVERLAY WINDOW
# ----------------------------------------------------------------------
class OverlayRenderer:
    def __init__(self, esp, window_size, menu):
        self.esp = esp
        self.window_size = window_size
        self.menu = menu
        self.scene = QtWidgets.QGraphicsScene()
        
        self.draw_box = True
        self.draw_health = True
        self.draw_names = True

        self.color_accent = QtGui.QColor("#ff1d65")
        self.color_white = QtGui.QColor(255, 255, 255, 255)
        self.color_bg = QtGui.QColor("#121212")

    def update(self):
        self.scene.clear()
        
        if not (self._is_window_focused(title="Counter-Strike 2") or self._is_window_focused(hwnd=int(self.menu.winId()))):
            return

        self._draw_watermark()
        self.esp.update_entities()

        for entity in self.esp.entities:
            if not entity.pos or entity.lifestate == 258:
                continue
            if entity.team == self.esp.local_player.team:
                continue
            self._draw_entity(entity)

    def _draw_watermark(self):
        current_time = time.strftime("%H:%M")
        text = f"| {current_time}"
        font = QtGui.QFont("Segoe UI", 10, QtGui.QFont.Bold)
        
        text_item = QtWidgets.QGraphicsTextItem(text)
        text_item.setFont(font)
        text_item.setDefaultTextColor(self.color_accent)
        th = text_item.boundingRect().height()
        tw = text_item.boundingRect().width()
        
        logo_path = get_resource_path("logo.svg")
        svg_item = QGraphicsSvgItem(logo_path)
        svg_rect = svg_item.boundingRect()
        
        target_size = 18.0
        scale_factor = target_size / svg_rect.height() if svg_rect.height() > 0 else 1.0
        svg_item.setScale(scale_factor)
        svg_w = svg_rect.width() * scale_factor

        padding_x = 12
        spacing = 6
        h = 32  
        w = (padding_x * 2) + svg_w + spacing + tw
        x, y = 15, 15
        
        path = QtGui.QPainterPath()
        path.addRoundedRect(x, y, w, h, 12, 12)
        self.scene.addPath(path, QtGui.QPen(self.color_accent, 1), QtGui.QBrush(self.color_bg))
        
        svg_item.setPos(x + padding_x, y + (h - target_size) / 2)
        self.scene.addItem(svg_item)
        
        text_y_offset = 1.5 
        text_item.setPos(x + padding_x + svg_w + spacing, y + (h - th) / 2 - text_y_offset)
        self.scene.addItem(text_item)

    def _draw_entity(self, entity):
        feet_pos = entity.pos
        if not feet_pos: return

        head_pos = (feet_pos[0], feet_pos[1], feet_pos[2] + 70.0) 

        head_screen = self._world_to_screen(head_pos)
        feet_screen = self._world_to_screen(feet_pos)
        if not head_screen or not feet_screen: return

        height = feet_screen[1] - head_screen[1]
        width = height / 2
        box_x = feet_screen[0] - width / 2
        box_y = head_screen[1]

        if self.draw_box:
            box = QtWidgets.QGraphicsRectItem(box_x, box_y, width, height)
            box.setPen(QtGui.QPen(self.color_white, 1.5))
            self.scene.addItem(box)

        if self.draw_health:
            health_height = height * (entity.health / 100)
            health_bar = QtWidgets.QGraphicsRectItem(box_x - 6, box_y + height - health_height, 2, health_height)
            health_bar.setBrush(QtGui.QBrush(self.color_accent))
            health_bar.setPen(QtGui.QPen(QtCore.Qt.NoPen))
            self.scene.addItem(health_bar)

        if self.draw_names:
            font = QtGui.QFont("Segoe UI", 8)
            name_item = self.scene.addText(entity.name, font)
            name_item.setDefaultTextColor(self.color_white)
            name_item.setPos(box_x + width / 2 - name_item.boundingRect().width() / 2, box_y - name_item.boundingRect().height())

    def _world_to_screen(self, pos):
        try:
            matrix = [self.esp.pm.read_float(self.esp.client + Offsets.dwViewMatrix + i * 4) for i in range(16)]
            
            x = matrix[0] * pos[0] + matrix[1] * pos[1] + matrix[2] * pos[2] + matrix[3]
            y = matrix[4] * pos[0] + matrix[5] * pos[1] + matrix[6] * pos[2] + matrix[7]
            w = matrix[12] * pos[0] + matrix[13] * pos[1] + matrix[14] * pos[2] + matrix[15]
            
            if w < 0.01: return None
                
            inv_w = 1.0 / w
            return (
                self.window_size[0] / 2 * (1 + x * inv_w),
                self.window_size[1] / 2 * (1 - y * inv_w)
            )
        except Exception:
            return None

    def _is_window_focused(self, title=None, hwnd=None):
        if title: hwnd = win32gui.FindWindow(None, title)
        elif hwnd is None: return False
        return hwnd and win32gui.GetForegroundWindow() == hwnd

class ESPOverlay(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        
        # 1. Configurar Audio Principal (Solo mantenemos el Click)
        self.audio_output = QAudioOutput()
        self.audio_output.setVolume(1.0)
        
        self.click_player = QMediaPlayer()
        self.click_player.setAudioOutput(self.audio_output)
        
        # 2. Iniciar Discord RPC
        self.init_rpc()

        self.esp = ESP()
        if not self.esp.initialize():
            sys.exit("[Overlay] Failed to initialize ESP")

        self.window_size = self._get_game_window_size()
        if not self.window_size[0]:
            sys.exit("[Overlay] Could not find CS2 window")

        self.menu = OverlayMenu(self)
        self.renderer = OverlayRenderer(self.esp, self.window_size, self.menu)
        
        self.menu_visible = False
        self.esp_enabled = True
        self._setup_overlay_window()
        self._setup_graphics()
        self._setup_timers()

    def init_rpc(self):
        if not HAS_RPC:
            print("[RPC] Librería pypresence no detectada.")
            return
            
        try:
            client_id = "1549528115032752258" 
            self.rpc = Presence(client_id)
            self.rpc.connect()
            
            self.rpc.update(
                details="PIGWARE",
                state="Free",
                large_image="logo",
                large_text=".gg/EkuMzj2Rk",
                buttons=[{"label": "WEBSITE", "url": "https://notcni.github.io/pigware/"}]
            )
            print("[RPC] Discord Rich Presence Conectado.")
        except Exception as e:
            print(f"[RPC] Error conectando a Discord: {e}")

    def _setup_overlay_window(self):
        self.setGeometry(0, 0, *self.window_size)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.Tool)
        win32gui.SetWindowLong(
            self.winId(), win32con.GWL_EXSTYLE,
            win32gui.GetWindowLong(self.winId(), win32con.GWL_EXSTYLE) | win32con.WS_EX_LAYERED | win32con.WS_EX_TRANSPARENT
        )

    def _setup_graphics(self):
        self.view = QtWidgets.QGraphicsView(self.renderer.scene, self)
        self.view.setGeometry(0, 0, *self.window_size)
        self.view.setStyleSheet("background: transparent; border: none;")
        self.view.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.view.setSceneRect(0, 0, *self.window_size)
        self.view.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.view.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        
        # Filtros Anti-Aliasing (Para bordes de watermark lisos sin dientes de sierra)
        self.view.setRenderHint(QtGui.QPainter.Antialiasing)
        self.view.setRenderHint(QtGui.QPainter.TextAntialiasing)
        self.view.setRenderHint(QtGui.QPainter.SmoothPixmapTransform)

    def _setup_timers(self):
        timer = QtCore.QTimer(self)
        timer.timeout.connect(self._update_overlay)
        timer.start(16)

        key_timer = QtCore.QTimer(self)
        key_timer.timeout.connect(self._check_insert_key)
        key_timer.start(100)

    def _update_overlay(self):
        if self.esp_enabled:
            self.renderer.update()
        else:
            self.renderer.scene.clear()

    def _check_insert_key(self):
        VK_INSERT = 0x2D
        if win32api.GetAsyncKeyState(VK_INSERT) & 0x8000:
            if self._is_window_focused(title="Counter-Strike 2") or self._is_window_focused(hwnd=int(self.menu.winId())):
                if not hasattr(self, "_insert_pressed") or not self._insert_pressed:
                    self._toggle_menu()
                    self._insert_pressed = True
        else:
            self._insert_pressed = False

    def _toggle_menu(self):
        self.menu_visible = not self.menu_visible

        if self.menu_visible:
            if not hasattr(self, "_menu_initialized_pos"):
                screen_width = self.window_size[0]
                menu_width = self.menu.width()
                self.menu.move(screen_width - menu_width - 40, 40)
                self._menu_initialized_pos = True

            self.menu.setWindowOpacity(0.0)
            self.menu.setVisible(True)
            self.anim = QtCore.QPropertyAnimation(self.menu, b"windowOpacity")
            self.anim.setDuration(250)
            self.anim.setStartValue(0.0)
            self.anim.setEndValue(1.0)
            self.anim.setEasingCurve(QtCore.QEasingCurve.OutCubic)
            self.anim.start()

            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('%')
            win32gui.SetForegroundWindow(int(self.menu.winId()))

            menu_geom = self.menu.frameGeometry()
            x = menu_geom.left() + menu_geom.width() // 2
            y = menu_geom.top() + menu_geom.height() // 2
            win32api.SetCursorPos((x, y))
        else:
            self.anim = QtCore.QPropertyAnimation(self.menu, b"windowOpacity")
            self.anim.setDuration(250)
            self.anim.setStartValue(1.0)
            self.anim.setEndValue(0.0)
            self.anim.setEasingCurve(QtCore.QEasingCurve.InCubic)
            self.anim.finished.connect(self.menu.hide)
            self.anim.start()
            
            hwnd_game = win32gui.FindWindow(None, "Counter-Strike 2")
            if hwnd_game:
                win32gui.SetForegroundWindow(hwnd_game)

    def _get_game_window_size(self):
        hwnd = win32gui.FindWindow(None, "Counter-Strike 2")
        if hwnd:
            rect = win32gui.GetWindowRect(hwnd)
            return (rect[2] - rect[0], rect[3] - rect[1])
        return (None, None)

    def _is_window_focused(self, title=None, hwnd=None):
        if title: hwnd = win32gui.FindWindow(None, title)
        elif hwnd is None: return False
        return hwnd and win32gui.GetForegroundWindow() == hwnd

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    
    # Asignar el icono a la aplicación general
    icon_path = get_resource_path("icon.ico")
    app.setWindowIcon(QtGui.QIcon(icon_path))
    
    overlay = ESPOverlay()
    overlay.show()
    sys.exit(app.exec())
