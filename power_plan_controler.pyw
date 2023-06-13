import ctypes
from PySide6 import QtCore, QtWidgets, QtGui

HIGH_PERFORMANCE_APP_MAP : set = {"firefox.exe"} 
BALANCED_APP_MAP         : set = {"chrome.exe", "Code.exe"} 

class GUID(ctypes.Structure):
    """
    guiddef.h
    https://learn.microsoft.com/ja-jp/windows/win32/api/guiddef/ns-guiddef-guid
    
    ---
    ``` C++
    typedef struct _GUID {
    unsigned long  Data1;
    unsigned short Data2;
    unsigned short Data3;
    unsigned char  Data4[8];
    } GUID;
    ```
    ---

    """
    _fields_ = [("Data1", ctypes.c_ulong),     # unsigned long  Data1;
                ("Data2", ctypes.c_ushort),    # unsigned short Data2;
                ("Data3", ctypes.c_ushort),    # unsigned short Data3;
                ("Data4", ctypes.c_ubyte * 8)] # unsigned char  Data4[8];

    def __str__(self) -> str:
        return '{:08x}-{:04x}-{:04x}-{}-{}'.format( 
        self.Data1,
        self.Data2,
        self.Data3,
        ''.join(f'{b:02x}' for b in self.Data4[:2]),
        ''.join(f'{b:02x}' for b in self.Data4[2:])
        )
    
    def __eq__(self, other) -> bool:
        if type(other) == GUID:
            return (
            self.Data1 == other.Data1 and
            self.Data2 == other.Data2 and
            self.Data3 == other.Data3 and 
            all(x == y for x, y in zip(self.Data4, other.Data4))
            )
        raise TypeError()

    def from_string(self,guid_str: str):
        data1, data2, data3, data4_1, data4_2 = guid_str.split('-')
        data1 = int(data1, 16)
        data2 = int(data2, 16)
        data3 = int(data3, 16)
        data4 = [int(data4_1[0:2], 16),int(data4_1[2:4], 16)] + [int(data4_2[i:i+2], 16) for i in range(0, len(data4_2), 2)]
        return GUID(
            ctypes.c_ulong(data1),
            ctypes.c_ushort(data2),
            ctypes.c_ushort(data3),
            (ctypes.c_ubyte*8)(*data4)
        )

class PowerPlanSetter():
    """
    電源オプションの設定を行う
    
    Attributes:
        POWER_SAVER      ( obj: GUID ) : 省電力

        BALANCED         ( obj: GUID ) : バランス

        HIGH_PERFORMANCE ( obj: GUID ) : 高パフォーマンス

    """
    
    POWER_SAVER : GUID       = GUID(ctypes.c_ulong(0xa1841308),
                            ctypes.c_ushort(0x3541),
                            ctypes.c_ushort(0x4fab),
                            (ctypes.c_ubyte*8)(0xbc,0x81,0xf7,0x15,0x56,0xf2,0x0b,0x4a)
                            )
    
    HIGH_PERFORMANCE : GUID = GUID(ctypes.c_ulong(0x8c5e7fda),
                            ctypes.c_ushort(0xe8bf),
                            ctypes.c_ushort(0x4a96),
                            (ctypes.c_ubyte*8)(0x9a,0x85,0xa6,0xe2,0x3a,0x8c,0x63,0x5c)
                            )
    
    BALANCED : GUID          = GUID(ctypes.c_ulong(0x381b4222),
                            ctypes.c_ushort(0xf694),
                            ctypes.c_ushort(0x41f0),
                            (ctypes.c_ubyte*8)(0x96,0x85,0xff,0x5b,0xb2,0x60,0xdf,0x2e)
                            )
    
    def __init__(self) -> None:
        self.power_prof_dll = ctypes.WinDLL("PowrProf")
        """ PowrProf.dll
        https://learn.microsoft.com/ja-jp/windows/win32/api/powrprof/
        """

    def get_active_power_plan(self) -> GUID:
        """
        現在アクティブな電源プランのGUIDを取得する。

        Args:
            なし

        Return: 
            GUID :
                現在アクティブな電源プランのGUIDを返す。
        """
        power_plan_p = ctypes.pointer(GUID())
        self.power_prof_dll.PowerGetActiveScheme(None, ctypes.byref(power_plan_p))
        return power_plan_p.contents
    
    def set_power_plan(self,power_plan : GUID) -> None:
        """
        電源プランを設定する。

        Ex: 高パフォーマンスに設定したい場合
            set_power_plan(powersetting.high_performance)

        Args:
            power_plan (GUID) : 設定する電源プランのGUID

        Return: なし
        """
        self.power_prof_dll.PowerSetActiveScheme(None,ctypes.byref(power_plan))

    def power_plan_str(self,power_plan : GUID) -> str:
        """
        電源プランのGUIDから、電源プランの名前を取得する。

        Args:
            GUID : 電源プランのGUID

        Return: 
            str : 電源プランの名前
        """
        if   power_plan == self.HIGH_PERFORMANCE:
            return "高パフォーマンス"
        elif power_plan == self.BALANCED:
            return "バランス"
        elif power_plan == self.POWER_SAVER:
            return "省電力"
        
        return ""
    

class app_list_s(ctypes.Structure):
    """
    "proc_list.h"
    
    ---
    ``` c
    typedef struct APP_LIST {
        size_t count;
        char** names;
    } app_list_s;
    ```
    ---
    """
    _fields_ = [("count", ctypes.c_size_t),
                ("names", ctypes.POINTER(ctypes.c_char_p))]

class ProcessListManager():
    def __init__(self):
        self.proc_list_dll = ctypes.cdll.LoadLibrary("./proc_list.dll")
        self.process_list = app_list_s()
        if self.proc_list_dll.new_app_list(ctypes.byref(self.process_list)) != 0:
            raise Exception("メモリの確保に失敗")
    def get(self) -> list:
        is_get_list_ok = self.proc_list_dll.get_process_list(ctypes.byref(self.process_list))
        if is_get_list_ok != 0 or not isinstance(self.process_list.count, int) or self.process_list.count <= 0:
            raise Exception("プロセス一覧の取得に失敗")
        proc_list = [self.process_list.names[i].decode('shift-jis') for i in range(self.process_list.count)]
        return proc_list
    def __del__(self):
        self.proc_list_dll.del_app_list(ctypes.byref(self.process_list))
        
class DynamicPowerPlanController(QtCore.QThread):
    
    """ DynamicPowerPlanController

    実行中のアプリケーションに応じて、適切な電源プランを設定する。
    
    Args:
        high_perf_apps : 高パフォーマンスアプリケーションのリスト  
        balanced_apps  : バランスアプリケーションのリスト

    """

    def __init__(self, high_perf_apps : set, balanced_apps : set) -> None:
        super().__init__()
        self.high_perf_apps : set = high_perf_apps
        self.balanced_apps :  set = balanced_apps
        self.power_plan_setter = PowerPlanSetter()
        self.request_stop = False
        self.process_list_manager = ProcessListManager()
        self.prev_tasklist: set = set()

    def set_power_plan_based_on_running_apps(self):
        tasklist : set = set(self.process_list_manager.get())
        if self.prev_tasklist == tasklist:
            return True
        self.prev_tasklist = tasklist
        if tasklist & self.high_perf_apps:
            self.power_plan_setter.set_power_plan(PowerPlanSetter.HIGH_PERFORMANCE)
            return True
        elif tasklist & self.balanced_apps:
            self.power_plan_setter.set_power_plan(PowerPlanSetter.BALANCED)
            return True
        else :
            self.power_plan_setter.set_power_plan(PowerPlanSetter.POWER_SAVER)
            return True

    def run(self):
        self.request_stop = False
        while not self.request_stop:
            self.set_power_plan_based_on_running_apps()
            for i in range(5):
                if self.request_stop :
                    break
                else:
                    self.msleep(500)

    def stop(self):
        self.request_stop = True
        if self.isRunning():
            self.wait()

class SysTray(QtWidgets.QSystemTrayIcon):
    def __init__(self,app : QtWidgets.QApplication ,power_plan_controller_thread :  DynamicPowerPlanController):
        super().__init__()
        self.q_app : QtWidgets.QApplication = app
        self.icon__ = QtGui.QIcon("icon.png")
        self.setIcon(self.icon__)

        self.power_plan_controller_thread : DynamicPowerPlanController = power_plan_controller_thread
        self.power_plan_setter = PowerPlanSetter()

        self.init_menu()
        
    def init_menu(self):
        # メニューの作成
        self.menu = QtWidgets.QMenu()

        self.auto_setter_action = QtGui.QAction('auto', self.menu)
        self.auto_setter_action.setObjectName('auto')
        self.auto_setter_action.triggered.connect(self.set_auto)  # type: ignore
        self.auto_setter_action.setCheckable(True)
        self.menu.addAction(self.auto_setter_action)
        self.menu.setDefaultAction(self.auto_setter_action)

        self.menu.addSeparator()

        self.high_performance_action = QtGui.QAction('high_performance', self.menu)
        self.high_performance_action.setObjectName('high_performance')
        self.high_performance_action.triggered.connect(self.set_high_performance)     # type: ignore
        self.high_performance_action.setCheckable(True)
        self.menu.addAction(self.high_performance_action)

        self.balanced_action = QtGui.QAction('balanced', self.menu)
        self.balanced_action.setObjectName('balanced')
        self.balanced_action.triggered.connect(self.set_balance)    # type: ignore
        self.balanced_action.setCheckable(True)
        self.menu.addAction(self.balanced_action)

        self.power_saver_action = QtGui.QAction('power_save', self.menu)
        self.power_saver_action.setObjectName('m_power_save')
        self.power_saver_action.triggered.connect(self.set_power_save)  # type: ignore
        self.power_saver_action.setCheckable(True)
        self.menu.addAction(self.power_saver_action)

        self.menu.addSeparator()

        self.exit_app_action = QtGui.QAction('exit', self.menu)
        self.exit_app_action.setObjectName('a_exit')
        self.exit_app_action.triggered.connect(self.exit_app)     # type: ignore
        self.exit_app_action.setCheckable(False)
        self.menu.addAction(self.exit_app_action)
        
        self.setContextMenu(self.menu)
        self.checked_action = None

    def switch_checked_action(self,new_action):
        if self.checked_action is not None:
            self.checked_action.setChecked(False)
        self.checked_action = new_action
        self.checked_action.setChecked(True)

    def set_auto(self):
        self.power_plan_controller_thread.start()
        self.switch_checked_action(self.auto_setter_action)

    def set_high_performance(self):
        self.power_plan_controller_thread.stop()
        self.power_plan_setter.set_power_plan(PowerPlanSetter.HIGH_PERFORMANCE)
        self.switch_checked_action(self.high_performance_action)

    def set_balance(self):
        self.power_plan_controller_thread.stop()
        self.power_plan_setter.set_power_plan(PowerPlanSetter.BALANCED)
        self.switch_checked_action(self.balanced_action)

    def set_power_save(self):
        self.power_plan_controller_thread.stop()
        self.power_plan_setter.set_power_plan(PowerPlanSetter.POWER_SAVER)
        self.switch_checked_action(self.power_saver_action)

    def exit_app(self):
        self.power_plan_controller_thread.stop()
        return self.q_app.quit()

def main():
    app = QtWidgets.QApplication([])
    power_plan_controller_thread = DynamicPowerPlanController(HIGH_PERFORMANCE_APP_MAP,BALANCED_APP_MAP)
    sys_tray_icon = SysTray(
        app,
        power_plan_controller_thread
    )
    sys_tray_icon.set_auto()
    sys_tray_icon.show()
    app.exec()


if __name__ == "__main__":
    main()
