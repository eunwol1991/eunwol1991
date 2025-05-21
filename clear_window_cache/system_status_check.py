import psutil
import platform
import datetime
import os

def bytes_to_gb(b):
    return round(b / (1024 ** 3), 2)

print("🧠 系统状态检测中...\n")

# 获取系统信息
uname = platform.uname()
print(f"💻 系统：{uname.system} {uname.release} ({uname.version})")
print(f"🖥️ 主机名：{uname.node}")
print(f"🧠 处理器：{uname.processor}")
print(f"⏰ 开机时间：{datetime.datetime.fromtimestamp(psutil.boot_time()).strftime('%Y-%m-%d %H:%M:%S')}")
print("\n==============================")

# CPU 使用率
print(f"🧮 CPU 使用率：{psutil.cpu_percent(interval=1)}%")

# 内存情况
mem = psutil.virtual_memory()
print(f"🧠 内存总量：{bytes_to_gb(mem.total)} GB")
print(f"📦 已使用：{bytes_to_gb(mem.used)} GB")
print(f"💤 可用：{bytes_to_gb(mem.available)} GB")
print(f"💢 使用率：{mem.percent}%")

# 磁盘情况
print("\n🗃️ 磁盘使用情况：")
for part in psutil.disk_partitions():
    try:
        usage = psutil.disk_usage(part.mountpoint)
        print(f"  - {part.device} [{part.mountpoint}]：{usage.percent}% 已使用（{bytes_to_gb(usage.used)} / {bytes_to_gb(usage.total)} GB）")
    except PermissionError:
        continue

# 活跃进程 TOP 5
print("\n🔥 占资源最多的前 5 个进程：")
top_procs = sorted(psutil.process_iter(['pid', 'name', 'cpu_percent', 'memory_percent']), 
                   key=lambda p: (p.info['cpu_percent'], p.info['memory_percent']), 
                   reverse=True)[:5]

for p in top_procs:
    info = p.info
    print(f"  - PID {info['pid']}: {info['name']} | CPU: {info['cpu_percent']}% | MEM: {round(info['memory_percent'], 2)}%")

print("\n✅ 状态检测完成！")
