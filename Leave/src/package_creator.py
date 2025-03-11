import subprocess
import os
import sys


def create_executable():
    """
    使用 PyInstaller 将项目打包成单文件可执行程序。
    模块会自动寻找 src/main.py 作为入口文件，模板文件位于 templates 目录下。
    """
    # 项目根目录（假设 package_creator.py 位于 src 目录下）
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    main_script = os.path.join(base_dir, 'src', 'main.py')
    template_file = os.path.join(base_dir, 'templates', '请假条模板.docx')

    if not os.path.isfile(main_script):
        print(f"主程序文件不存在: {main_script}")
        return

    if not os.path.isfile(template_file):
        print(f"模板文件不存在: {template_file}")
        return

    # 输出目录（项目根目录下的 dist 目录）
    output_dir = os.path.join(base_dir, "dist")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 系统分隔符，Windows 使用 ';'，其他使用 ':'
    separator = ';' if sys.platform == 'win32' else ':'

    pyinstaller_command = [
        'pyinstaller',
        '--onefile',
        '--windowed',
        '--hidden-import=pandas',
        '--hidden-import=python-docx',
        '--add-data', f'{template_file}{separator}templates',
        '--distpath', output_dir,
        '--workpath', os.path.join(output_dir, 'build'),
        '--specpath', os.path.join(output_dir, 'spec'),
        main_script
    ]

    try:
        subprocess.run(pyinstaller_command, check=True)
        print(f"打包成功：{main_script}")
    except subprocess.CalledProcessError as e:
        print(f"打包失败: {e}")
        print("标准输出:", e.stdout)
        print("标准错误:", e.stderr)


if __name__ == "__main__":
    create_executable()
