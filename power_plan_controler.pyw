import ctypes

import subprocess
import time

class GUID(ctypes.Structure):
    _fields_ = [("Data1", ctypes.c_ulong),
                ("Data2", ctypes.c_ushort),
                ("Data3", ctypes.c_ushort),
                ("Data4", ctypes.c_ubyte * 8)]
    
    def string(self) -> str:
        return '{:08x}-{:04x}-{:04x}-{}-{}'.format( 
        self.Data1,
        self.Data2,
        self.Data3,
        ''.join(f'{b:02x}' for b in self.Data4[:2]),
        ''.join(f'{b:02x}' for b in self.Data4[2:])
        )
    
    def from_string(guid_str: str):
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
    
    def __eq__(self, other: object) -> bool:
        if type(other) == GUID:
            return (
            self.Data1 == other.Data1 and
            self.Data2 == other.Data2 and
            self.Data3 == other.Data3 and 
            all(x == y for x, y in zip(self.Data4, other.Data4))
            )
        raise TypeError()

class powersetting():
    """
    電源オプションの設定を行う
    """
    
    
    power_save : GUID       = GUID(ctypes.c_ulong(0xa1841308),
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
    
    
    balance : GUID          = GUID(ctypes.c_ulong(0x381b4222),
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
        elif power_plan == self.balance:
            return "バランス"
        elif power_plan == self.power_save:
            return "省電力"
        
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

def process_exists(processes : list):
    tasklist = subprocess.run('tasklist', shell=True, capture_output=True, text=True)
    for process in processes:
        if process in tasklist.stdout:
            return True
    return False

def main_loop(high_performance_process):
    power = powersetting()
    is_high_performance = power.get_active_power_plan() == powersetting.high_performance
    while True:
        if process_exists(high_performance_process):
            if not is_high_performance:
                power.set_power_plan(powersetting.high_performance)
                is_high_performance = True
        else:
            if is_high_performance:
                power.set_power_plan(powersetting.power_save)
                is_high_performance = False
        time.sleep(5)

def main():
    # 高パフォーマンスで実行したいプロセス名を入れる
    high_performance_process : list = ["r5apex.exe"]
    main_loop(high_performance_process)


main()
