"""兼容旧脚本接口的包装器。"""

from excel2word.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
