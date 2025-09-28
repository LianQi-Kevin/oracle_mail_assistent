import aria2p
from concurrent.futures import ThreadPoolExecutor, wait, ALL_COMPLETED
import time
from tqdm import tqdm


DOWNLOAD_PATH = r"./downloads"

# aria2p
RPC_PORT: int = 12768    # aria2c RPC 端口
RPC_SECRET: str = ""     # aria2c RPC 密钥（留空则不使用密钥）
ARIA2P_API = aria2p.API(aria2p.Client(host="http://localhost", port=RPC_PORT, secret=RPC_SECRET))

# gids
DOWNLOAD_GIDS: list[str] = []


def download_options(download_path: str, file_name: str) -> dict:
    return {
        "dir": download_path,  # 下载目录
        "out": file_name,  # 保存的文件名
        "continue": "true",  # 断点续传
        "max-connection-per-server": "16",  # 每个服务器的最大连接数
        "split": 16,  # 文件分片数
        "min-split-size": "1M",  # 最小分片大小
        "file-allocation": "falloc",  # 文件预分配方式
        }


def download_progress_shower():
    global DOWNLOAD_GIDS, ARIA2P_API
    # 初始化总进度条（任务数）和当前任务进度条（字节进度）
    total_bar = tqdm(total=0, desc="总进度 0/0", unit="任务", position=0)
    current_bar = None
    current_gid = None
    done_count = 0
    tracked = []

    while True:
        time.sleep(0.5)  # 避免循环过快
        # 检测并添加新任务
        for gid in DOWNLOAD_GIDS:
            if gid not in tracked:
                tracked.append(gid)
                total_bar.total += 1
                total_bar.set_description(f"总进度 {done_count}/{total_bar.total}")

        # 找到下一个未完成的任务作为当前任务
        for gid in tracked:
            download = ARIA2P_API.get_download(gid)
            if download.status == "complete":
                # 如果任务完成且还未统计，则更新完成计数并关闭进度条
                if gid != current_gid:
                    done_count += 1
                    total_bar.update(1)
                    total_bar.set_description(f"总进度 {done_count}/{total_bar.total}")
                if gid == current_gid and current_bar is not None:
                    current_bar.close()
                    current_bar = None
                    current_gid = None
                continue

            # 如果当前任务进度条不是这个 gid，则切换到新任务
            if current_gid != gid:
                if current_bar is not None:
                    current_bar.close()
                total_size = int(download.total_length or 0)
                current_bar = tqdm(
                    total=total_size,
                    desc=f"任务 {done_count+1} / 大小 {total_size}B",
                    unit="B", unit_scale=True, unit_divisor=1024,
                    position=1
                )
                current_gid = gid
                current_downloaded = 0
            # 更新当前任务进度条
            completed = int(download.completed_length or 0)
            delta = completed - current_bar.n
            if delta > 0:
                current_bar.update(delta)
            # 显示下载速度（MB/s）
            speed = download.download_speed or 0
            current_bar.set_postfix_str(f"{speed/1024/1024:.1f} MB/s")
            current_bar.refresh()
            break

        # 如果所有任务完成，退出循环
        if done_count == total_bar.total and total_bar.total > 0:
            break

    # 关闭进度条
    total_bar.close()
    if current_bar is not None:
        current_bar.close()


if __name__ == '__main__':
    pool = ThreadPoolExecutor(max_workers=1)
    future = pool.submit(download_progress_shower)

    urls: dict[str, str] = {
        "may-cad-64.exe": "https://www.framexpert.com/media/download/windows/maytec/may-cad-64.exe",
        "pycharm-2025.2.2.exe": "https://download-cdn.jetbrains.com/python/pycharm-2025.2.2.exe"
    }

    for file_name, url in urls.items():
        download = ARIA2P_API.add(url, options=download_options(download_path=DOWNLOAD_PATH, file_name=file_name))
        DOWNLOAD_GIDS.append(download[0].gid)
        print(f"gid: {download[0].gid}, 保存到 {DOWNLOAD_PATH}/{file_name}")

    # waiting mission complete
    wait([future], return_when=ALL_COMPLETED)
    pool.shutdown()
