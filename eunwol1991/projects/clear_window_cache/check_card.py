import subprocess

def check_nvidia_usage():
    try:
        result = subprocess.check_output("nvidia-smi", shell=True, encoding="utf-8")
        print("🎮 检测到 NVIDIA 独立显卡使用中：")
        print("="*50)
        print(result)
    except Exception as e:
        print("❌ 没有检测到 NVIDIA 独立显卡或驱动未安装")
        print(f"错误详情：{e}")

if __name__ == "__main__":
    check_nvidia_usage()
