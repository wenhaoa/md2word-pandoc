"""Retired one-off migration. Importing or running this file never writes files."""
import sys


def main():
    print('旧版一次性迁移脚本已停用；请直接维护SKILL.md并运行回归测试。', file=sys.stderr)
    return 2


if __name__ == '__main__':
    sys.exit(main())
