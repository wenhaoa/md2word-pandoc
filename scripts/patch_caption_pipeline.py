"""Retired one-off patch. Importing or running this file never writes files."""
import sys


def main():
    print('旧版一次性补丁脚本已停用；当前管线无需重复打补丁。', file=sys.stderr)
    return 2


if __name__ == '__main__':
    sys.exit(main())
