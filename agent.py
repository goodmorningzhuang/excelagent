"""
Excel处理Agent - 本地运行，无需API Key（交互式命令行版本）

运行后显示可选技能列表，选择技能后输入文件名进行处理。
支持一次性输入多个文件名批量处理。
"""

import os
import sys
import importlib.util
from datetime import datetime

# 项目根目录
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
PENDING_DIR = os.path.join(ROOT_DIR, "pending")
SKILL_DIR = os.path.join(ROOT_DIR, "skill")

# ANSI 颜色代码
RED = "\033[91m"
GREEN = "\033[92m"
YELLOW = "\033[93m"
BLUE = "\033[94m"
CYAN = "\033[96m"
BOLD = "\033[1m"
RESET = "\033[0m"


def print_banner():
    print()
    print(f"{CYAN}{'=' * 60}{RESET}")
    print(f"{CYAN}{BOLD}  Excel处理Agent - 本地运行{RESET}")
    print(f"{CYAN}{'=' * 60}{RESET}")
    print()


def discover_skills():
    skills = {}
    if not os.path.exists(SKILL_DIR):
        print(f"{RED}错误: skill/ 目录不存在{RESET}")
        return skills

    for item in os.listdir(SKILL_DIR):
        skill_path = os.path.join(SKILL_DIR, item)
        if os.path.isdir(skill_path):
            if os.path.exists(os.path.join(skill_path, "__init__.py")) or \
               os.path.exists(os.path.join(skill_path, "handler.py")):
                skill_info = {
                    "name": item,
                    "path": skill_path,
                    "module": None,
                    "description": ""
                }
                desc_file = os.path.join(skill_path, "skill.txt")
                if os.path.exists(desc_file):
                    with open(desc_file, "r", encoding="utf-8") as f:
                        skill_info["description"] = f.read().strip()
                skills[item] = skill_info

    return skills


def load_skill_module(skill_name):
    handler_path = os.path.join(SKILL_DIR, skill_name, "handler.py")
    if not os.path.exists(handler_path):
        print(f"  {RED}错误: 技能 {skill_name} 缺少 handler.py{RESET}")
        return None

    module_name = f"skill_{skill_name}"
    spec = importlib.util.spec_from_file_location(module_name, handler_path)
    if spec is None or spec.loader is None:
        print(f"  {RED}错误: 无法加载技能模块 {handler_path}{RESET}")
        return None

    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as e:
        print(f"  {RED}错误: 技能模块执行失败 - {e}{RESET}")
        return None

    return module


def generate_output_path(input_path, skill_name):
    """
    根据输入路径生成输出路径。
    原文件名 + 时间戳 + after + .xlsx
    例如: 测试文件.xls -> 测试文件_20250101_120000_after.xlsx
    """
    directory = os.path.dirname(input_path)
    basename = os.path.basename(input_path)
    name_without_ext = os.path.splitext(basename)[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_name = f"{name_without_ext}_{timestamp}_after.xlsx"
    output_path = os.path.join(directory, output_name)
    return output_path


def list_skill_files(skill_name):
    pending_folder = os.path.join(PENDING_DIR, skill_name)
    if not os.path.exists(pending_folder):
        return []

    files = []
    for item in sorted(os.listdir(pending_folder)):
        item_path = os.path.join(pending_folder, item)
        if os.path.isdir(item_path):
            continue
        if not (item.endswith('.xls') or item.endswith('.xlsx')):
            continue
        if item.startswith('~$'):
            continue
        files.append(item)
    return files


def show_skill_menu(skills):
    skill_list = list(skills.keys())
    print(f"{BOLD}可用技能列表:{RESET}")
    print(f"{CYAN}{'-' * 50}{RESET}")

    for i, name in enumerate(skill_list, 1):
        info = skills[name]
        desc = info['description']
        first_line = desc.split('\n')[0] if desc else "无描述"
        print(f"  {GREEN}[{i}]{RESET} {BOLD}{name}{RESET}  - {first_line}")

    print(f"{CYAN}{'-' * 50}{RESET}")
    print()

    while True:
        try:
            choice = input(f"{YELLOW}请选择技能 (输入数字，或 q 退出): {RESET}").strip()
            if choice.lower() == 'q' or choice.lower() == 'quit':
                return None
            idx = int(choice) - 1
            if 0 <= idx < len(skill_list):
                return skill_list[idx]
            else:
                print(f"  {RED}无效选择，请输入 1-{len(skill_list)} 之间的数字{RESET}")
        except ValueError:
            print(f"  {RED}请输入数字{RESET}")
        except (EOFError, KeyboardInterrupt):
            return None


def select_files_interactive(skill_name):
    files = list_skill_files(skill_name)

    if not files:
        pending_folder = os.path.join(PENDING_DIR, skill_name)
        print(f"  {YELLOW}pending/{skill_name}/ 目录下没有找到Excel文件{RESET}")
        print(f"  {YELLOW}请将待处理文件放入: {pending_folder}{RESET}")
        return []

    print(f"\n{BOLD}pending/{skill_name}/ 目录下的Excel文件:{RESET}")
    print(f"{CYAN}{'-' * 50}{RESET}")

    for i, fname in enumerate(files, 1):
        print(f"  {GREEN}[{i}]{RESET} {fname}")

    print(f"  {GREEN}[0]{RESET} 处理全部文件")
    print(f"{CYAN}{'-' * 50}{RESET}")
    print()

    pending_folder = os.path.join(PENDING_DIR, skill_name)
    selected = []

    while True:
        try:
            choice = input(f"{YELLOW}选择文件 (输入序号，多个用空格/逗号分隔，q 返回): {RESET}").strip()
            if choice.lower() == 'q' or choice.lower() == 'quit':
                return None

            if choice == '0':
                selected = [os.path.join(pending_folder, f) for f in files]
                break

            choice = choice.replace(',', ' ')
            indices = choice.split()
            valid = True
            temp_selected = []
            for idx_str in indices:
                try:
                    idx = int(idx_str) - 1
                    if 0 <= idx < len(files):
                        temp_selected.append(os.path.join(pending_folder, files[idx]))
                    else:
                        print(f"  {RED}无效序号: {idx_str}，请输入 1-{len(files)}{RESET}")
                        valid = False
                        break
                except ValueError:
                    print(f"  {RED}无效输入: {idx_str}，请输入数字{RESET}")
                    valid = False
                    break
            if valid and temp_selected:
                selected = temp_selected
                break
        except (EOFError, KeyboardInterrupt):
            return None

    return selected


def process_files(skill_name, skill_info, file_paths):
    print(f"\n  正在加载技能模块...")
    module = load_skill_module(skill_name)
    if module is None:
        print(f"  {RED}错误: 无法加载技能模块{RESET}")
        return 0

    if not hasattr(module, 'process'):
        print(f"  {RED}错误: 技能模块缺少 process 函数{RESET}")
        return 0

    print(f"  {GREEN}技能模块加载成功{RESET}")

    processed_count = 0
    total = len(file_paths)

    print(f"\n{BOLD}开始处理文件 ({total} 个):{RESET}")
    print(f"{CYAN}{'-' * 50}{RESET}")

    for i, file_path in enumerate(file_paths, 1):
        filename = os.path.basename(file_path)
        print(f"\n  [{i}/{total}] 处理: {filename}")

        if not os.path.exists(file_path):
            print(f"    {RED}文件不存在，跳过{RESET}")
            continue

        try:
            output_path = generate_output_path(file_path, skill_name)
            result_path = module.process(file_path, output_path)
            if result_path:
                processed_count += 1
                print(f"    {GREEN}✓ 处理成功{RESET}")
            else:
                processed_count += 1
                print(f"    {GREEN}✓ 处理完成{RESET}")
        except Exception as e:
            print(f"    {RED}✗ 处理失败: {e}{RESET}")

    print(f"\n{CYAN}{'-' * 50}{RESET}")
    print(f"  {BOLD}处理完成! 成功 {GREEN}{processed_count}{RESET}/{total} 个文件{RESET}")

    return processed_count


def skill_submenu(skills, skill_name):
    skill_info = skills[skill_name]

    print(f"\n{CYAN}{'=' * 60}{RESET}")
    print(f"{BOLD}  技能: {skill_name}{RESET}")
    print(f"{CYAN}{'=' * 60}{RESET}")

    desc_lines = skill_info['description'].split('\n')
    print(f"\n{BOLD}技能功能:{RESET}")
    for line in desc_lines[:6]:
        print(f"  {line}")

    while True:
        print(f"\n{BOLD}操作选择:{RESET}")
        print(f"  {GREEN}[1]{RESET} 从文件列表中选择")
        print(f"  {GREEN}[q]{RESET} 返回上级菜单")
        print()

        try:
            choice = input(f"{YELLOW}请选择: {RESET}").strip()
        except (EOFError, KeyboardInterrupt):
            return None

        if choice == '1':
            file_paths = select_files_interactive(skill_name)
            if file_paths is None:
                continue
            if not file_paths:
                continue
            process_files(skill_name, skill_info, file_paths)
            break

        elif choice.lower() == 'q':
            return None


def main():
    while True:
        print_banner()

        if not os.path.exists(PENDING_DIR):
            os.makedirs(PENDING_DIR)
            print(f"  {YELLOW}已创建 {PENDING_DIR}/ 目录{RESET}")

        print(f"{BOLD}正在扫描技能...{RESET}")
        skills = discover_skills()
        print(f"  共发现 {GREEN}{len(skills)}{RESET} 个技能")
        print()

        if not skills:
            print(f"  {RED}没有找到可用技能，请在 skill/ 目录下添加技能模块{RESET}")
            input(f"\n按回车键退出...")
            return

        selected_skill = show_skill_menu(skills)

        if selected_skill is None:
            print(f"\n{YELLOW}已退出，感谢使用!{RESET}\n")
            break

        skill_submenu(skills, selected_skill)

        print()
        try:
            again = input(f"{YELLOW}按回车键继续，输入 q 退出: {RESET}").strip()
            if again.lower() == 'q':
                print(f"\n{YELLOW}已退出，感谢使用!{RESET}\n")
                break
        except (EOFError, KeyboardInterrupt):
            print(f"\n{YELLOW}已退出，感谢使用!{RESET}\n")
            break

        print("\n" * 3)


if __name__ == "__main__":
    main()