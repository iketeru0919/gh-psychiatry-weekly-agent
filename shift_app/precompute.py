"""全施設のサマリを事前計算する(夜間バッチ用)。

使い方:
    python -m shift_app.precompute [データDIR]

cron 例(毎日 5:00):
    0 5 * * * cd /path/to/app && python -m shift_app.precompute
"""
import sys
import time

from .model_manager import ModelManager
from .store import Store
from .web import DATA_ROOT


def main():
    root = sys.argv[1] if len(sys.argv) > 1 else DATA_ROOT
    store = Store(root)
    manager = ModelManager(store)
    facs = store.facilities()
    print(f'{len(facs)} 施設のサマリを計算します')
    for f in facs:
        t0 = time.time()
        try:
            manager.summary(f['id'], force=True)
            print(f"  OK {f['name']} ({f['year']}/{f['month']}) {time.time()-t0:.1f}s")
        except Exception as e:
            print(f"  NG {f['name']}: {e}")
        manager.drop(f['id'])  # メモリ解放
    print('完了')


if __name__ == '__main__':
    main()
