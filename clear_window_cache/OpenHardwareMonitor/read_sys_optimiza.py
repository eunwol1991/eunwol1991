import clr
import sys
import os
import time

# 设置 DLL 搜索路径（当前目录）
dll_dir = os.path.abspath(".")
sys.path.append(dll_dir)

# 加载 OpenHardwareMonitorLib.dll（不含扩展名）
clr.AddReference("OpenHardwareMonitorLib")

# 导入 .NET 类
from OpenHardwareMonitor import Hardware

# 初始化监控器
computer = Hardware.Computer()
computer.CPUEnabled = True
computer.Open()

def get_cpu_temperatures():
    temps = []
    for hw in computer.Hardware:
        hw.Update()
        if hw.HardwareType == Hardware.HardwareType.CPU:
            # 多次刷新以避免读取为 None
            for _ in range(3):
                hw.Update()
                time.sleep(0.1)
            for sensor in hw.Sensors:
                if sensor.SensorType == Hardware.SensorType.Temperature:
                    name = sensor.Name
                    value = sensor.Value
                    temps.append((name, value))
    return temps

# 示例：读取并打印一次
if __name__ == "__main__":
    print("🌡️ 当前 CPU 温度：")
    for name, temp in get_cpu_temperatures():
        temp_display = f"{temp:.1f} °C" if temp is not None else "N/A"
        print(f"  {name}: {temp_display}")
