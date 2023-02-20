import ctypes
from PySide6 import QtCore, QtWidgets, QtGui

# 注意!! : 要素数が1個のタプルを生成する場合は、末尾にカンマ,が必要です。
HIGH_PERFORMANCE_APP_MAP : tuple = ( "r5apex.exe", )
BALANCED_APP_MAP         : tuple = ( "firefox.exe", "Chrome.exe", "Code.exe" )

class GUID(ctypes.Structure):
    _fields_ = [("Data1", ctypes.c_ulong),
                ("Data2", ctypes.c_ushort),
                ("Data3", ctypes.c_ushort),
                ("Data4", ctypes.c_ubyte * 8)]

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
    """
    
    POWER_SAVER : GUID       = GUID(ctypes.c_ulong(0xa1841308),
                            ctypes.c_ushort(0x3541),
                            ctypes.c_ushort(0x4fab),
                            (ctypes.c_ubyte*8)(0xbc,0x81,0xf7,0x15,0x56,0xf2,0x0b,0x4a)
                            )
    """ 電源プラン : 省電力 """
    
    HIGH_PERFORMANCE : GUID = GUID(ctypes.c_ulong(0x8c5e7fda),
                            ctypes.c_ushort(0xe8bf),
                            ctypes.c_ushort(0x4a96),
                            (ctypes.c_ubyte*8)(0x9a,0x85,0xa6,0xe2,0x3a,0x8c,0x63,0x5c)
                            )
    """ 電源プラン : 高パフォーマンス """
    
    BALANCED : GUID          = GUID(ctypes.c_ulong(0x381b4222),
                            ctypes.c_ushort(0xf694),
                            ctypes.c_ushort(0x41f0),
                            (ctypes.c_ubyte*8)(0x96,0x85,0xff,0x5b,0xb2,0x60,0xdf,0x2e)
                            )
    """電源プラン : バランス"""
    
    def __init__(self) -> None:
        self.power_prof_dll = ctypes.WinDLL("PowrProf")

    def get_active_power_plan(self) -> GUID:
        """
        現在アクティブな電源プランのGUIDを取得する。

        引数:
            なし

        返り値: 
            GUID :
                現在アクティブな電源プランのGUIDを返す。
        """
        power_plan_p = ctypes.pointer(GUID())
        self.power_prof_dll.PowerGetActiveScheme(None, ctypes.byref(power_plan_p))
        return power_plan_p.contents
    
    def set_power_plan(self,power_plan : GUID) -> None:
        """
        電源プランを設定する。

            例: 高パフォーマンスに設定したい場合
                set_power_plan(powersetting.high_performance)

        引数:
            power_plan (GUID) :
                設定する電源プランのGUID

        返り値: 
            なし
        """
        self.power_prof_dll.PowerSetActiveScheme(None,ctypes.byref(power_plan))

    def power_plan_str(self,power_plan : GUID) -> str:
        """
        ### 説明
            電源プランのGUIDから、電源プランの名前を取得する。

        ### 引数
            GUID : 電源プランのGUID

        ### 返り値
            str : 電源プランの名前
        """
        if power_plan == self.HIGH_PERFORMANCE:
            return "高パフォーマンス"
        elif power_plan == self.BALANCED:
            return "バランス"
        elif power_plan == self.POWER_SAVER:
            return "省電力"
        
        return ""

class app_list_s(ctypes.Structure):
    _fields_ = [("count", ctypes.c_size_t),
                ("names", ctypes.POINTER(ctypes.c_char_p))]
class ProcessListManager():
    def __init__(self):
        self.proc_list_dll = ctypes.cdll.LoadLibrary("./proc_list.dll")
        self.process_list = app_list_s()
    def get(self) -> list:
        if self.proc_list_dll.new_app_list(ctypes.byref(self.process_list)) != 0:
            return [None]
        if self.proc_list_dll.get_process_list(ctypes.byref(self.process_list)) != 0 :
            return [None]
        if self.process_list.count is ctypes.c_int or self.process_list.count <= 0 :
            return [None]
        proc_list = [self.process_list.names[i].decode('shift-jis') for i in range(int(self.process_list.count))]
        self.proc_list_dll.del_app_list(ctypes.byref(self.process_list))
        return proc_list
        


class DynamicPowerPlanController(QtCore.QThread):
    
    """
        実行中のアプリケーションに応じて、適切な電源プランを設定する。
        
        引数:

            high_perf_apps - 高パフォーマンスアプリケーションのリスト  

            balanced_apps - バランスアプリケーションのリスト

    """

    def __init__(self, high_perf_apps : tuple, balanced_apps : tuple) -> None:
        super().__init__()
        self.high_perf_apps : tuple = high_perf_apps
        self.balanced_apps :  tuple = balanced_apps
        self.power_plan_setter = PowerPlanSetter()
        self.request_stop = False
        self.process_list_manager = ProcessListManager()


    def set_power_plan_based_on_running_apps(self):
        tasklist = self.process_list_manager.get()
        # if high_performance_apps in processes -> set high_performance
        for app in self.high_perf_apps:
            if app in tasklist:
                self.power_plan_setter.set_power_plan(PowerPlanSetter.HIGH_PERFORMANCE)
                return True
        
        # else if balanced_apps in processes    -> set balanced
        for app in self.balanced_apps:
            if app in tasklist:
                self.power_plan_setter.set_power_plan(PowerPlanSetter.BALANCED)
                return True

        # else -> set power_saver
        self.power_plan_setter.set_power_plan(PowerPlanSetter.POWER_SAVER)
        return True

    def run(self):
        self.request_stop = False
        counter = 0
        while not self.request_stop:
            if counter == 5 :
                self.set_power_plan_based_on_running_apps()
                counter = 0
            else:
                self.sleep(1)
                counter += 1

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

        self.checked_action : QtGui.QAction = self.auto_setter_action
        self.auto_setter_action.setChecked(True)
        self.power_plan_controller_thread.start()
        
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

    def switch_checked_action(self,new_action):
        self.checked_action.setChecked(False)
        self.checked_action = new_action
        self.checked_action.setChecked(True)

    def set_auto(self):
        self.switch_checked_action(self.auto_setter_action)
        self.power_plan_controller_thread.start()

    def set_high_performance(self):
        self.switch_checked_action(self.high_performance_action)
        self.power_plan_controller_thread.stop()
        self.power_plan_setter.set_power_plan(PowerPlanSetter.HIGH_PERFORMANCE)

    def set_balance(self):
        self.switch_checked_action(self.balanced_action)
        self.power_plan_controller_thread.stop()
        self.power_plan_setter.set_power_plan(PowerPlanSetter.BALANCED)

    def set_power_save(self):
        self.switch_checked_action(self.power_saver_action)
        self.power_plan_controller_thread.stop()
        self.power_plan_setter.set_power_plan(PowerPlanSetter.POWER_SAVER)

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
    sys_tray_icon.show()
    app.exec()
    
main()
