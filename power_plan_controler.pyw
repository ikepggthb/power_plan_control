import ctypes
import subprocess
import time

from PySide6 import QtCore, QtWidgets, QtGui

high_performance_app_map : list = ["r5apex"]
balanced_app_map : list = ["firefox","Chrome","Code"]

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

class powersetting():
    """
    電源オプションの設定を行う
    """
    
    power_saver : GUID       = GUID(ctypes.c_ulong(0xa1841308),
                            ctypes.c_ushort(0x3541),
                            ctypes.c_ushort(0x4fab),
                            (ctypes.c_ubyte*8)(0xbc,0x81,0xf7,0x15,0x56,0xf2,0x0b,0x4a)
                            )
    """ 電源プラン : 省電力 """
    
    high_performance : GUID = GUID(ctypes.c_ulong(0x8c5e7fda),
                            ctypes.c_ushort(0xe8bf),
                            ctypes.c_ushort(0x4a96),
                            (ctypes.c_ubyte*8)(0x9a,0x85,0xa6,0xe2,0x3a,0x8c,0x63,0x5c)
                            )
    """ 電源プラン : 高パフォーマンス """
    
    balanced : GUID          = GUID(ctypes.c_ulong(0x381b4222),
                            ctypes.c_ushort(0xf694),
                            ctypes.c_ushort(0x41f0),
                            (ctypes.c_ubyte*8)(0x96,0x85,0xff,0x5b,0xb2,0x60,0xdf,0x2e)
                            )
    """電源プラン : バランス"""
    
    def __init__(self) -> None:
        self.powersetting = ctypes.WinDLL("PowrProf")

    def get_active_power_plan(self) -> GUID:
        """
        現在アクティブな電源プランのGUIDを取得する。

        引数:
            なし

        返り値: 
            GUID :
                現在アクティブな電源プランのGUIDを返す。
        """
        poweroption_p = ctypes.pointer(GUID())
        self.powersetting.PowerGetActiveScheme(None, ctypes.byref(poweroption_p))
        return poweroption_p.contents
    
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
        self.powersetting.PowerSetActiveScheme(None,ctypes.byref(power_plan))

    def power_plan_str(self,power_plan) -> str:
        """
        ### 説明
            電源プランのGUIDから、電源プランの名前を取得する。

        ### 引数
            GUID : 電源プランのGUID

        ### 返り値
            str : 電源プランの名前
        """
        if power_plan == self.high_performance:
            return "高パフォーマンス"
        elif power_plan == self.balanced:
            return "バランス"
        elif power_plan == self.power_saver:
            return "省電力"
        
        return ""
        
    def get_active_power_plan_name(self):
        """
        ### 説明
            現在アクティブな電源プランの名前を取得する。

        ### 引数
            なし

        ### 返り値
            str : 電源プランの名前
        """
        return self.power_plan_str(self.get_active_power_plan())


class powerplan_set_thread(QtCore.QThread):
    def __init__(self,power : powersetting) -> None:
        super().__init__()
        self.power : powersetting = power
        self.is_should_loop_stop = False
        self.is_loop_stop = True

    def powerplan_set_loop(self) -> None :
        self.is_should_loop_stop = False
        self.is_loop_stop = False
        while not self.is_should_loop_stop:
            self.powerplan_set()
            time.sleep(3)
        self.is_loop_stop = True

    def powerplan_set(self):
        tasklist = subprocess.run('tasklist', shell=True, capture_output=True, text=True).stdout

        # if high_performance_app in processes -> set high_performance
        for app in high_performance_app_map:
            if app in tasklist:
                self.power.set_power_plan(powersetting.high_performance)
                return True
        
        # else if balanced_app in processes    -> set balanced
        for app in balanced_app_map:
            if app in tasklist:
                self.power.set_power_plan(powersetting.balanced)
                return True
        
        # else -> set power_saver
        self.power.set_power_plan(powersetting.power_saver)
        return True
    
    def run(self):
        self.powerplan_set_loop()

    def stop(self):
        self.is_should_loop_stop = True

class systray(QtWidgets.QSystemTrayIcon):
    def __init__(self,app : QtWidgets.QApplication ,pp_thread :  powerplan_set_thread, power : powersetting):
        super().__init__()
        self.app = app
        self.icon__ = QtGui.QIcon("icon.png")
        self.setIcon(self.icon__)

        self.pp_thread = pp_thread
        self.power = power

        self.init_menu()

        self.checked_action = self.a_auto
        self.a_auto.setChecked(True)
        self.pp_thread.start()
        
    def init_menu(self):
        # メニューの作成
        self.menu = QtWidgets.QMenu()

        self.a_auto = QtGui.QAction('auto', self.menu)
        self.a_auto.setObjectName('m_auto')
        self.a_auto.triggered.connect(self.set_auto)  # type: ignore
        self.a_auto.setCheckable(True)
        self.menu.addAction(self.a_auto)
        self.menu.setDefaultAction(self.a_auto)

        # ------------------------
        self.menu.addSeparator()

        self.a_high_performance = QtGui.QAction('high_performance', self.menu)
        self.a_high_performance.setObjectName('m_high_performance')
        self.a_high_performance.triggered.connect(self.set_high_performance)     # type: ignore
        self.a_high_performance.setCheckable(True)
        self.menu.addAction(self.a_high_performance)

        self.a_balanced = QtGui.QAction('balanced', self.menu)
        self.a_balanced.setObjectName('m_balanced')
        self.a_balanced.triggered.connect(self.set_balance)    # type: ignore
        self.a_balanced.setCheckable(True)
        self.menu.addAction(self.a_balanced)

        self.a_power_saver = QtGui.QAction('power_save', self.menu)
        self.a_power_saver.setObjectName('m_power_save')
        self.a_power_saver.triggered.connect(self.set_power_save)  # type: ignore
        self.a_power_saver.setCheckable(True)
        self.menu.addAction(self.a_power_saver)

        # ------------------------
        self.menu.addSeparator()

        # 終了アクション
        self.a_exit = QtGui.QAction('exit', self.menu)
        self.a_exit.setObjectName('a_exit')
        self.a_exit.triggered.connect(self.exit_app)     # type: ignore
        self.a_exit.setCheckable(False)
        self.menu.addAction(self.a_exit)
        
        self.setContextMenu(self.menu)

    def move_checked_action(self,action):
        self.checked_action.setChecked(False)
        self.checked_action = action
        self.checked_action.setChecked(True)

    def set_auto(self):
        self.move_checked_action(self.a_auto)
        self.pp_thread.start()

    def set_high_performance(self):
        self.move_checked_action(self.a_high_performance)
        self.pp_thread.stop()
        self.power.set_power_plan(powersetting.high_performance)

    def set_balance(self):
        self.move_checked_action(self.a_balanced)
        self.pp_thread.stop()
        self.power.set_power_plan(powersetting.balanced)

    def set_power_save(self):
        self.move_checked_action(self.a_power_saver)
        self.pp_thread.stop()
        self.power.set_power_plan(powersetting.power_saver)

    def exit_app(self):
        self.pp_thread.stop()
        # powerplan_set_threadのループが終わるまで待つ。
        while not self.pp_thread.is_loop_stop:
            time.sleep(0.1)
        return self.app.quit()

def main():
    app = QtWidgets.QApplication([])
    power = powersetting()
    pp_thread = powerplan_set_thread(power)
    sys_tray_icon = systray(app,pp_thread,power)
    sys_tray_icon.show()
    app.exec()
    
main()
