import time
import subprocess
import sys


def compare():
    print("СРАВНЕНИЕ ВРЕМЕНИ ВЫПОЛНЕНИЯ")
    print("-" * 35)

    # запуск синхронной версии
    start = time.time()
    subprocess.run([sys.executable, "пандас.py", "async_data"])
    sync_time = time.time() - start

    # запуск асинхронной версии
    start = time.time()
    subprocess.run([sys.executable, "асинхронка2.py", "async_data"])
    async_time = time.time() - start

    # результат
    print("\n" + "=" * 35)
    print("РЕЗУЛЬТАТЫ:")
    print(f"  Синхронный:  {sync_time:.3f} сек")
    print(f"  Асинхронный: {async_time:.3f} сек")
    print(f"  Ускорение:   {sync_time / async_time:.2f}x")
    print("=" * 35)


if __name__ == "__main__":
    compare()