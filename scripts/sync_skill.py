#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Skill 与项目代码的双向同步脚本

Usage:
    python scripts/sync_skill.py pull   # 从 Skill 拉取到项目（Skill 作为镜像源）
    python scripts/sync_skill.py push   # 从项目推送到 Skill
    python scripts/sync_skill.py diff   # 显示差异
"""
import argparse
import filecmp
import os
import shutil
import sys
from pathlib import Path


# 技能目录（镜像源）
SKILL_ROOT = Path(__file__).resolve().parent.parent / "skills" / "strategic-emerging-monthly-report"

# 项目根目录
PROJECT_ROOT = Path(__file__).resolve().parent.parent

# 映射表：skill 路径 → 项目路径（相对路径）
PULL_MAPPINGS = [
    ("src/csv_word_converter/", "src/csv_word_converter/"),
    ("scripts/", "scripts/"),
    ("references/", "references/"),
    ("ab_response_formats/", "ab_response_formats/"),
    ("agents/", "agents/"),
    ("templates_config.yaml", "templates_config.yaml"),
    ("requirements.txt", "requirements.txt"),
]

# 推送时的反向映射
PUSH_MAPPINGS = [
    ("src/csv_word_converter/", "src/csv_word_converter/"),
    ("scripts/", "scripts/"),
    ("references/", "references/"),
    ("ab_response_formats/", "ab_response_formats/"),
    ("agents/", "agents/"),
    ("templates_config.yaml", "templates_config.yaml"),
    ("requirements.txt", "requirements.txt"),
]


def ensure_dir(path: Path) -> None:
    """确保目录存在，Windows 兼容"""
    path.mkdir(parents=True, exist_ok=True)


def copy_file(src: Path, dst: Path, verbose: bool = True) -> bool:
    """复制单个文件，支持 Windows 路径"""
    try:
        ensure_dir(dst.parent)
        shutil.copy2(src, dst)
        if verbose:
            print(f"  ✓ {src} -> {dst}")
        return True
    except Exception as e:
        print(f"  ✗ {src} -> {dst}: {e}", file=sys.stderr)
        return False


def copy_dir(src: Path, dst: Path, verbose: bool = True) -> bool:
    """复制整个目录，支持 Windows 路径"""
    try:
        if dst.exists():
            shutil.rmtree(dst)
        ensure_dir(dst.parent)
        shutil.copytree(src, dst)
        if verbose:
            print(f"  ✓ {src}/ -> {dst}/")
        return True
    except Exception as e:
        print(f"  ✗ {src}/ -> {dst}/: {e}", file=sys.stderr)
        return False


def sync_mappings(mappings: list, skill_root: Path, project_root: Path, verbose: bool = True) -> int:
    """执行映射同步，返回成功复制的文件/目录数"""
    success_count = 0
    for src_rel, dst_rel in mappings:
        src = skill_root / src_rel
        dst = project_root / dst_rel

        if src.is_dir():
            if src.exists():
                if copy_dir(src, dst, verbose):
                    success_count += 1
            elif verbose:
                print(f"  ⚠ 跳过不存在的目录: {src}")
        elif src.is_file():
            if src.exists():
                if copy_file(src, dst, verbose):
                    success_count += 1
            elif verbose:
                print(f"  ⚠ 跳过不存在的文件: {src}")
        else:
            if verbose:
                print(f"  ⚠ 跳过不存在的路径: {src}")
    return success_count


def compare_mappings(mappings: list, skill_root: Path, project_root: Path) -> int:
    """比较映射并显示差异，返回有差异的项数"""
    diff_count = 0

    for src_rel, dst_rel in mappings:
        src = skill_root / src_rel
        dst = project_root / dst_rel

        if src.is_dir() and dst.is_dir():
            # 比较两个目录
            comparison = filecmp.dircmp(str(src), str(dst))
            diffs = _collect_diffs(comparison)
            if diffs:
                print(f"\n📁 {src_rel} ↔ {dst_rel}:")
                for diff in diffs:
                    print(f"  {diff}")
                diff_count += len(diffs)
            else:
                print(f"  ✓ {src_rel} 与 {dst_rel} 完全相同")
        elif src.exists() and dst.exists():
            # 比较两个文件
            if not filecmp.cmp(src, dst, shallow=False):
                print(f"\n📄 {src_rel} ≠ {dst_rel}")
                diff_count += 1
            else:
                print(f"  ✓ {src_rel} 与 {dst_rel} 相同")
        elif src.exists() and not dst.exists():
            print(f"\n🆕 仅在 Skill 存在: {src_rel}")
            diff_count += 1
        elif not src.exists() and dst.exists():
            print(f"\n🆕 仅在项目存在: {dst_rel}")
            diff_count += 1
        else:
            print(f"  - 两者均不存在: {src_rel}")

    return diff_count


def _collect_diffs(dcmp: filecmp.dircmp, prefix: str = "") -> list:
    """递归收集目录比较中的所有差异"""
    diffs = []

    # 文件差异
    for name in dcmp.diff_files:
        diffs.append(f"  {prefix}修改: {name}")

    # 仅在左侧（skill）存在的文件/目录
    for name in dcmp.left_only:
        path = os.path.join(dcmp.left, name)
        if os.path.isdir(path):
            diffs.append(f"  {prefix}Skill 独有目录: {name}")
        else:
            diffs.append(f"  {prefix}Skill 独有文件: {name}")

    # 仅在右侧（project）存在的文件/目录
    for name in dcmp.right_only:
        path = os.path.join(dcmp.right, name)
        if os.path.isdir(path):
            diffs.append(f"  {prefix}项目独有目录: {name}")
        else:
            diffs.append(f"  {prefix}项目独有文件: {name}")

    # 递归比较子目录
    for sub_dcmp in dcmp.subdirs.values():
        sub_prefix = prefix + "  "
        diffs.extend(_collect_diffs(sub_dcmp, sub_prefix))

    return diffs


def pull(args: argparse.Namespace) -> int:
    """从 Skill 拉取到项目（Skill 作为镜像源）"""
    print("=" * 60)
    print("🔄 从 Skill 拉取到项目")
    print("=" * 60)
    print(f"Skill 目录: {SKILL_ROOT}")
    print(f"项目目录:   {PROJECT_ROOT}")
    print()

    if not SKILL_ROOT.exists():
        print(f"❌ Skill 目录不存在: {SKILL_ROOT}", file=sys.stderr)
        return 1

    count = sync_mappings(PULL_MAPPINGS, SKILL_ROOT, PROJECT_ROOT)
    print()
    print(f"✅ 完成！成功同步 {count} 项")

    if args.git_add:
        _run_git_add()

    return 0


def push(args: argparse.Namespace) -> int:
    """从项目推送到 Skill"""
    print("=" * 60)
    print("🔄 从项目推送到 Skill")
    print("=" * 60)
    print(f"项目目录:   {PROJECT_ROOT}")
    print(f"Skill 目录: {SKILL_ROOT}")
    print()

    if not SKILL_ROOT.exists():
        print(f"❌ Skill 目录不存在: {SKILL_ROOT}", file=sys.stderr)
        return 1

    count = sync_mappings(PUSH_MAPPINGS, SKILL_ROOT, PROJECT_ROOT)
    print()
    print(f"✅ 完成！成功同步 {count} 项")

    if args.git_add:
        _run_git_add()

    return 0


def diff(args: argparse.Namespace) -> int:
    """显示 Skill 与项目之间的差异"""
    print("=" * 60)
    print("📊 Skill 与项目差异比较")
    print("=" * 60)
    print(f"Skill 目录: {SKILL_ROOT}")
    print(f"项目目录:   {PROJECT_ROOT}")
    print()

    if not SKILL_ROOT.exists():
        print(f"❌ Skill 目录不存在: {SKILL_ROOT}", file=sys.stderr)
        return 1

    diff_count = compare_mappings(PULL_MAPPINGS, SKILL_ROOT, PROJECT_ROOT)
    print()
    if diff_count > 0:
        print(f"⚠️  发现 {diff_count} 处差异")
    else:
        print("✅ Skill 与项目完全同步")

    return 0 if diff_count == 0 else 1


def _run_git_add() -> None:
    """运行 git add 命令（由 hook 调用时使用）"""
    try:
        import subprocess
        skill_path = SKILL_ROOT
        result = subprocess.run(
            ["git", "add", str(skill_path)],
            capture_output=True,
            text=True,
            cwd=PROJECT_ROOT
        )
        if result.returncode == 0:
            print(f"✅ 已执行 git add {skill_path}")
        else:
            print(f"⚠️  git add 失败: {result.stderr}")
    except Exception as e:
        print(f"⚠️  无法执行 git add: {e}")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Skill ↔ 项目代码双向同步工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python scripts/sync_skill.py pull          # 从 Skill 拉取到项目
  python scripts/sync_skill.py push          # 从项目推送到 Skill
  python scripts/sync_skill.py diff          # 显示差异
  python scripts/sync_skill.py pull --git-add  # 同步后执行 git add
        """
    )

    subparsers = parser.add_subparsers(dest="command", help="子命令")

    # pull 命令
    pull_parser = subparsers.add_parser("pull", help="从 Skill 拉取到项目")
    pull_parser.add_argument(
        "--git-add", action="store_true",
        help="同步后执行 git add"
    )
    pull_parser.set_defaults(func=pull)

    # push 命令
    push_parser = subparsers.add_parser("push", help="从项目推送到 Skill")
    push_parser.add_argument(
        "--git-add", action="store_true",
        help="同步后执行 git add"
    )
    push_parser.set_defaults(func=push)

    # diff 命令
    diff_parser = subparsers.add_parser("diff", help="显示 Skill 与项目差异")
    diff_parser.set_defaults(func=diff)

    args = parser.parse_args()

    if args.command is None:
        parser.print_help()
        return 1

    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
